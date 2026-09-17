import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';
import plan from './fixtures/word-pdf-whitespace/reference.json' with { type: 'json' };
const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
for (const sample of plan.rows.filter(row => row.kind === 'nbsp')) test(`Whitespace ${sample.name} font failure preserves editing and recovers NBSP export`, async ({ page }) => {
  const root = `.local/word-decoration-spaces/failure-browser-v3/${sample.name}`;
  await fs.mkdir(root, { recursive: true });
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(sample.path);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true }); await waitWordLayout(editor);
  const model = () => editor.evaluate(el => (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getJSON());
  const before = await model(); let downloads = 0;
  page.on('download', () => downloads++);
  await page.evaluate(() => {
    const target = window as any; target.whitespaceOriginalFonts = target.queryLocalFonts;
    target.queryLocalFonts = async () => [];
  });
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('regular font Arial is unavailable');
  expect(downloads).toBe(0); expect(await model()).toEqual(before);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  await page.evaluate(() => { const target = window as any; target.queryLocalFonts = target.whitespaceOriginalFonts; });
  await editor.focus(); await page.keyboard.press('Control+End'); await page.keyboard.type('X');
  const edited = await model(); expect(edited).not.toEqual(before);
  await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(before);
  await page.keyboard.press('Control+y'); await expect.poll(model).toEqual(edited);
  await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(before);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
  const path = `${root}/original.docx`; await (await pending).saveAs(path);
  expect(hash(await fs.readFile(path))).toBe(sample.sourceHash);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const title = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: title, exact: true }).click(); await waitWordLayout(editor);
  expect(await model()).toEqual(before);
  await page.getByRole('button', { name: 'Export', exact: true }).click(); const recovered = page.waitForEvent('download');
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await (await recovered).saveAs(`${root}/recovered.pdf`);
  await fs.writeFile(`${root}/report.json`, JSON.stringify({ name: sample.name, originalHash: hash(await fs.readFile(path)),
    recoveredPdfHash: hash(await fs.readFile(`${root}/recovered.pdf`)),
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-pdf-whitespace-failure.spec.ts')),
    scope: 'Missing-font failure emits no file and preserves the model, typing/history and exact original DOCX. Reload restores font access and exports the original NBSP document.' }, null, 2));
});
