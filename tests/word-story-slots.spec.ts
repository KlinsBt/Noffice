import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

test('Word independently edits all first/even/default header and footer slots', async ({
  page,
}) => {
  test.setTimeout(90000);
  await page.context().grantPermissions(['local-fonts']);
  const root = '.local/word-side-stories/slots/all';
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-side-stories/first-even.docx');
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'All story slots.docx',
    buffer: source,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  await expect(page.locator('.section-page')).toHaveCount(4);
  for (const kind of ['header', 'footer'])
    for (const slot of ['default', 'first', 'even']) {
      const label = slot === 'first' ? 'First-page' : slot === 'even' ? 'Even-page' : 'Default';
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      await page.getByRole('button', { name: `Section 1 — ${label} ${kind}`, exact: true }).click();
      const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
      await editor.focus();
      await page.keyboard.press('Control+Home');
      await page.keyboard.press('Shift+End');
      await page.keyboard.insertText(`Changed ${slot} ${kind}`);
      await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
      await expect(page.locator('.section-page')).toHaveCount(4);
      for (const [index, role] of ['first', 'even', 'default', 'even'].entries())
        if (role === slot)
          await expect(
            page
              .locator('.section-page')
              .nth(index)
              .locator('.word-page-story')
              .nth(kind === 'header' ? 0 : 1),
          ).toHaveText(`Changed ${slot} ${kind}`);
    }
  const hashes: Record<string, string> = {
    source: createHash('sha256').update(source).digest('hex'),
  };
  const capture = async (stage: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const waiting = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    await (await waiting).saveAs(`${root}/${stage}.docx`);
    hashes[stage] = createHash('sha256')
      .update(await fs.readFile(`${root}/${stage}.docx`))
      .digest('hex');
  };
  for (let i = 0; i < 6; i++) await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await capture('undo');
  expect(hashes.undo).toBe(hashes.source);
  for (let i = 0; i < 6; i++) await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await capture('edited');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const waiting = page.waitForEvent('download');
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await (await waiting).saveAs(`${root}/download.pdf`);
  hashes.pdf = createHash('sha256')
    .update(await fs.readFile(`${root}/download.pdf`))
    .digest('hex');
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'All story slots', exact: true }).click();
  await capture('reloaded');
  await page.locator('input[type=file][multiple]').setInputFiles(`${root}/reloaded.docx`);
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'reloaded',
  );
  await expect(page.locator('.section-page')).toHaveCount(4);
  for (const [index, slot] of ['first', 'even', 'default', 'even'].entries())
    await expect(page.locator('.section-page').nth(index).locator('.word-page-story')).toHaveText([
      `Changed ${slot} header`,
      `Changed ${slot} footer`,
    ]);
  expect(errors).toEqual([]);
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify({ hashes, errors, buildHash: await wordStoryBuildHash() }, null, 2),
  );
});

test('Word disposes header measurement views when replacing or closing a document', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto('/');
  for (const [name, count] of [
    ['first-even', 6],
    ['default', 2],
  ] as const) {
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles(`tests/fixtures/word-side-stories/${name}.docx`);
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
    await expect(page.locator('.section-page')).toHaveCount(4);
    await expect(page.locator('.word-story-measure-host')).toHaveCount(count);
    await page.getByRole('textbox', { name: 'Document text', exact: true }).focus();
  }
  await page.getByRole('button', { name: 'Choose a mode', exact: true }).click();
  await expect(page.locator('.word-story-measure-host')).toHaveCount(0);
  await page.getByRole('button', { name: 'Start Word', exact: true }).click();
  await page.getByRole('textbox', { name: 'Document text', exact: true }).fill('New body');
  await expect(page.locator('.word-story-measure-host')).toHaveCount(0);
  expect(errors).toEqual([]);
});
