import { test, expect, type Locator, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/word-nonbreaking-spaces/reference.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
async function lines(paragraph: Locator) {
  return paragraph.evaluate(p => {
    const output: { top: number; text: string }[] = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    let node: Node | null;
    while ((node = walker.nextNode())) {
      const text = node as Text; if (text.parentElement?.closest('.ProseMirror-widget,[contenteditable=false]')) continue;
      for (let i = 0; i < text.length; i++) {
        const range = document.createRange(); range.setStart(text, i); range.setEnd(text, i + 1);
        const top = range.getBoundingClientRect().top; let line = output.find(line => Math.abs(line.top - top) < .25);
        if (!line) { line = { top, text: '' }; output.push(line); }
        line.text += text.data[i];
      }
    }
    return output.map(line => line.text);
  });
}
async function capture(page: Page, root: string, stage: string, hashes: Record<string, string>) {
  for (const [extension, label] of [['docx', 'DOCX file Editable in Microsoft Word'], ['pdf', 'PDF file']]) {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    const path = `${root}/${stage}.${extension}`; await (await pending).saveAs(path); hashes[`${stage}.${extension}`] = hash(await fs.readFile(path));
  }
}
for (const sample of reference.rows) test(`NBSP ${sample.name}: replace gap, wrap, history, reload and actual files`, async ({ page }) => {
  test.setTimeout(120000);
  const run = process.env.NOFFICE_PDF_NBSP_RUN || 'browser-v3';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid NBSP evidence run');
  const root = `.local/word-nonbreaking-spaces/${run}/${sample.name}`; await fs.mkdir(root, { recursive: true });
  const hashes: Record<string, string> = {}, errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  expect(hash(await fs.readFile(sample.path))).toBe(sample.sourceHash);
  await page.context().grantPermissions(['local-fonts']); await page.setViewportSize({ width: 1500, height: 1400 });
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(sample.path);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true }); await waitWordLayout(editor);
  const target = editor.locator('p').nth(9), original = ['A ', 'AA\u00a0BB'], edited = ['A AA ', 'BB'];
  await expect.poll(() => lines(target)).toEqual(original); await capture(page, root, 'source', hashes);
  expect(hashes['source.docx']).toBe(sample.sourceHash);
  await editor.focus(); await page.keyboard.press('Control+Home');
  const steps = sample.controls.slice(0,9).reduce((n, p) => n + p.text.length + 1, 0) + 4;
  for (let i = 0; i < steps; i++) await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  expect(await editor.evaluate(el => {
    const editor = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    return editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to, '');
  })).toBe('\u00a0');
  await page.keyboard.press('Space'); await waitWordLayout(editor); await expect.poll(() => lines(target)).toEqual(edited);
  await page.keyboard.press('Control+z'); await waitWordLayout(editor); await expect.poll(() => lines(target)).toEqual(original);
  await page.keyboard.press('Control+y'); await waitWordLayout(editor); await expect.poll(() => lines(target)).toEqual(edited);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible(); await capture(page, root, 'edited', hashes);
  const title = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: title, exact: true }).click(); await waitWordLayout(editor);
  await expect.poll(() => lines(target)).toEqual(edited); await capture(page, root, 'reloaded', hashes);
  await page.locator('input[type=file][multiple]').setInputFiles(`${root}/edited.docx`); await waitWordLayout(editor);
  await expect.poll(() => lines(target)).toEqual(edited); await capture(page, root, 'reimported', hashes);
  expect(errors).toEqual([]);
  await fs.writeFile(`${root}/report.json`, JSON.stringify({ name: sample.name, hashes, sourceHash: sample.sourceHash,
    errors, buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-pdf-nonbreaking-spaces.spec.ts')),
    scope: 'Actual keyboard replacement of NBSP with a space changes wrapping; undo/redo, persistence and DOCX/PDF downloads/reimport. Independent native comparison is separate.' }, null, 2));
});
