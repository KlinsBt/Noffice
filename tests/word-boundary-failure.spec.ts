import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';
import plan from './fixtures/word-boundary-delete/browser-cases.json' with { type: 'json' };
const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
for (const sample of plan.rows.filter(row => ['paragraph-period', 'section-period', 'layout-paragraph-indent', 'wrap-arial-hanging'].includes(row.name))) test(`Boundary ${sample.name} font failure preserves history, originals and recovered export`, async ({ page }) => {
  const run = process.env.NOFFICE_BOUNDARY_FAILURE_RUN || 'failure-browser-v1';
  if (!/^failure-browser-v\d+$/.test(run)) throw Error('Invalid failure evidence run');
  const root = `.local/word-boundary-delete/${run}/${sample.name}`;
  await fs.mkdir(root, { recursive: true });
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(sample.fixture);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true }); await waitWordLayout(editor);
  const model = () => editor.evaluate(el => (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getJSON());
  const original = await model();
  await editor.evaluate(el => { const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    const from = e.state.doc.firstChild!.nodeSize - 1; e.commands.setTextSelection({ from, to: from + 2 }); e.view.focus(); });
  await page.keyboard.press('Delete'); await waitWordLayout(editor);
  const before = await model(); expect(before).not.toEqual(original); let downloads = 0;
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
  await editor.focus(); await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(original);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
  const path = `${root}/original.docx`; await (await pending).saveAs(path);
  expect(hash(await fs.readFile(path))).toBe(sample.sourceHash);
  await editor.focus(); await page.keyboard.press('Control+y'); await expect.poll(model).toEqual(before);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const title = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: title, exact: true }).click(); await waitWordLayout(editor);
  expect(await model()).toEqual(before);
  await page.getByRole('button', { name: 'Export', exact: true }).click(); const recovered = page.waitForEvent('download');
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await (await recovered).saveAs(`${root}/recovered.pdf`);
  await page.getByRole('button', { name: 'Export', exact: true }).click(); const editable = page.waitForEvent('download');
  await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
  await (await editable).saveAs(`${root}/recovered.docx`);
  await fs.writeFile(`${root}/report.json`, JSON.stringify({ name: sample.name, originalHash: hash(await fs.readFile(path)),
    recoveredDocxHash: hash(await fs.readFile(`${root}/recovered.docx`)),
    recoveredPdfHash: hash(await fs.readFile(`${root}/recovered.pdf`)),
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-boundary-failure.spec.ts')),
    scope: 'Missing-font failure emits no file and preserves the joined model; typing/history and boundary undo recover exact original DOCX bytes. Redo/reload preserve the joined model and recovered DOCX/PDF exports.' }, null, 2));
});
