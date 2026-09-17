import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { waitWordLayout, wordPlacements } from './word-pagination-helpers';

test('Word default paragraphs retain their fonts and paginate through actual exports and reload', async ({
  page,
}) => {
  const root = '.local/word-default-pagination';
  await fs.mkdir(root, { recursive: true });
  const external: string[] = [];
  page.on('request', (request) => {
    if (
      new URL(request.url()).origin !== 'http://127.0.0.1:4183' &&
      !request.url().startsWith('data:')
    )
      external.push(request.url());
  });
  await page.setViewportSize({ width: 1500, height: 1200 });
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Word', exact: true }).click();
  await page.evaluate(async () => {
    await navigator.serviceWorker.ready;
  });
  await expect.poll(() => page.evaluate(() => !!navigator.serviceWorker.controller)).toBe(true);
  await page.context().setOffline(true);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(page.getByRole('spinbutton', { name: 'Font size', exact: true })).toHaveValue('12');
  await expect(page.getByRole('combobox', { name: 'Font family', exact: true })).toHaveValue(
    'Inter',
  );
  const outputs: Record<string, string> = {};
  const download = async (stage: string, kind: 'docx' | 'pdf') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', {
        name: kind === 'pdf' ? 'PDF file' : 'DOCX file Editable in Microsoft Word',
        exact: true,
      })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${stage}.${kind}`, bytes);
    outputs[`${stage}.${kind}`] = createHash('sha256').update(bytes).digest('hex');
  };
  await download('empty', 'docx');
  await download('empty', 'pdf');
  await editor.focus();
  await page.keyboard.insertText('line01');
  for (let i = 2; i <= 70; i++) {
    await page.keyboard.press('Shift+Enter');
    await page.keyboard.insertText(`line${String(i).padStart(2, '0')}`);
  }
  const capture = async (stage: string) => {
    await waitWordLayout(editor);
    const points = await wordPlacements(editor);
    expect(points.flat().every((c) => c.page > 0)).toBe(true);
    await expect(page.locator('.section-page')).not.toHaveCount(0);
    await download(stage, 'docx');
    await download(stage, 'pdf');
    return points;
  };
  const source = await capture('source');
  await editor.focus();
  await page.keyboard.press('Control+End');
  await page.keyboard.insertText(' added');
  const edited = await capture('edited');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor).not.toContainText(' added');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(editor).toContainText(' added');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Untitled document', exact: true }).click();
  const restored = await capture('restored');
  expect(restored).toEqual(source);
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Default reimport.docx',
    buffer: await fs.readFile(`${root}/restored.docx`),
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Default reimport',
  );
  const reimported = await capture('reimported');
  expect(reimported).toEqual(source);
  expect(external).toEqual([]);
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        source,
        edited,
        restored,
        reimported,
        outputs,
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ),
  );
});
