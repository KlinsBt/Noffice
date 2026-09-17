import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout, wordLineOrigins } from './word-pagination-helpers';

async function open(page: Page, kind: string) {
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: `Section 1 — Default ${kind}`, exact: true }).click();
  return page.getByRole('textbox', { name: 'Header or footer text', exact: true });
}
async function docx(page: Page, path: string) {
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const waiting = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
    .click();
  await (await waiting).saveAs(path);
  return createHash('sha256')
    .update(await fs.readFile(path))
    .digest('hex');
}
for (const kind of ['header', 'footer'])
  test(`Word ${kind} growth repaginates with history, reload and actual exports`, async ({
    page,
  }) => {
    test.setTimeout(90000);
    await page.context().grantPermissions(['local-fonts']);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const root = `.local/word-side-stories/flow/${kind}`;
    await fs.mkdir(root, { recursive: true });
    await page.goto('/');
    const source = await fs.readFile('tests/fixtures/word-side-stories/default.docx');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `Grow ${kind}.docx`,
      buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(page.locator('.section-page')).toHaveCount(4);
    const story = await open(page, kind);
    await story.focus();
    await page.keyboard.press('Control+End');
    for (let i = 1; i <= 3; i++) {
      await page.keyboard.press('Shift+Enter');
      await page.keyboard.insertText(`extra${i}`);
    }
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await expect(page.locator('.section-page')).toHaveCount(5);
    await waitWordLayout(body);
    const initialLines = (await wordLineOrigins(body)).flat();
    expect(
      Array.from({ length: 5 }, (_, i) => initialLines.filter((l) => l.page === i + 1).length),
    ).toEqual([9, 9, 9, 9, 4]);
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(page.locator('.section-page')).toHaveCount(4);
    const hashes: Record<string, string> = {
      source: createHash('sha256').update(source).digest('hex'),
      undo: await docx(page, `${root}/undo.docx`),
    };
    expect(hashes.undo).toBe(hashes.source);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(page.locator('.section-page')).toHaveCount(5);
    hashes.edited = await docx(page, `${root}/edited.docx`);
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: 'PDF file', exact: true }).click();
    await (await waiting).saveAs(`${root}/download.pdf`);
    hashes.pdf = createHash('sha256')
      .update(await fs.readFile(`${root}/download.pdf`))
      .digest('hex');
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: `Grow ${kind}`, exact: true }).click();
    await expect(page.locator('.section-page')).toHaveCount(5);
    hashes.reloaded = await docx(page, `${root}/reloaded.docx`);
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/reloaded.docx`);
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'reloaded',
    );
    await expect(page.locator('.section-page')).toHaveCount(5);
    await waitWordLayout(body);
    expect((await wordLineOrigins(body)).flat().map((l) => l.page)).toEqual(
      initialLines.map((l) => l.page),
    );
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify(
        { hashes, errors, lines: initialLines, buildHash: await wordStoryBuildHash() },
        null,
        2,
      ),
    );
  });

test('Word rejects unsupported story table insertion, keeps a backup, and recovers with Undo', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto('/');
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles('tests/fixtures/word-side-stories/default.docx');
  const story = await open(page, 'header');
  await story.focus();
  await page.keyboard.press('Control+End');
  await page.keyboard.press('Enter');
  await story.evaluate((element) => {
    const clipboardData = new DataTransfer();
    clipboardData.setData('text/html', '<table><tr><td><p>New paragraph</p></td></tr></table>');
    element.dispatchEvent(
      new ClipboardEvent('paste', { clipboardData, bubbles: true, cancelable: true }),
    );
  });
  await expect(story.locator('table')).toHaveCount(1);
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  const downloads: string[] = [];
  page.on('download', (d) => downloads.push(d.suggestedFilename()));
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page
    .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
    .click();
  await expect(page.getByRole('alert')).toContainText('header/footer structural editing');
  expect(downloads).toEqual([]);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  await fs.mkdir('.local/word-side-stories/recovery', { recursive: true });
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const backupDownload = page.waitForEvent('download');
  await page.getByRole('button', { name: /Noffice backup Full editable model/ }).click();
  await (await backupDownload).saveAs('.local/word-side-stories/recovery/unsupported.noffice');
  const backup = JSON.parse(
    await fs.readFile('.local/word-side-stories/recovery/unsupported.noffice', 'utf8'),
  );
  expect(
    backup.content.stories.parts.some((p: { html: string }) => p.html.includes('New paragraph')),
  ).toBe(true);
  expect(
    Buffer.from(backup.original.base64, 'base64').equals(
      await fs.readFile('tests/fixtures/word-side-stories/default.docx'),
    ),
  ).toBe(true);
  await open(page, 'header');
  await expect(story).toContainText('New paragraph');
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.locator('.section-page')).toHaveCount(4);
  const hash = await docx(page, '.local/word-side-stories/recovery/restored.docx');
  expect(hash).toBe(
    createHash('sha256')
      .update(await fs.readFile('tests/fixtures/word-side-stories/default.docx'))
      .digest('hex'),
  );
  expect(errors).toEqual([]);
});
