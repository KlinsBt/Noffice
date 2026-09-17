import { test, expect, type Locator, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/word-pdf-tab-leader-decorations/reference.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout } from './word-pagination-helpers';
import JSZip from 'jszip';

const run = process.env.NOFFICE_PDF_TAB_LEADER_DECORATIONS_RUN || 'browser-v1';
if (!/^browser-(?:before-)?v\d+$/.test(run)) throw Error('Invalid PDF decoration run');
const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
async function marks(editor: Locator) {
  return editor.locator('p').first().evaluate(p => {
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), result: number[] = [];
    while (walker.nextNode()) {
      const text = walker.currentNode as Text;
      result.push(...Array.from(text.data, () => (text.parentElement?.closest('u') ? 1 : 0)
        | (text.parentElement?.closest('s,strike,del') ? 2 : 0)));
    }
    return result.slice(1, 2);
  });
}
async function capture(page: Page, root: string, stage: string, hashes: Record<string, string>) {
  for (const [extension, label] of [['docx', 'DOCX file Editable in Microsoft Word'], ['pdf', 'PDF file']]) {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    const path = `${root}/${stage}.${extension}`;
    await (await pending).saveAs(path); hashes[`${stage}.${extension}`] = hash(await fs.readFile(path));
  }
}
for (const sample of reference.rows) test(`Word PDF ${sample.name} formatting, history, reload and actual files`, async ({ page }) => {
  test.setTimeout(120000);
  const root = `.local/word-pdf-tab-leader-decorations/${run}/${sample.name}`;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile(sample.path); expect(hash(source)).toBe(sample.sourceHash);
  const hashes: Record<string, string> = {}, errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(sample.path);
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  await waitWordLayout(body); await capture(page, root, 'source', hashes);
  await body.screenshot({ path: `${root}/editor-source.png` });
  expect(hashes['source.docx']).toBe(sample.sourceHash);
  let editor = body;
  if (sample.context === 'stories') {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: /^Section 1 .*Default header/ }).click();
    editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    await editor.screenshot({ path: `${root}/story-editor-source.png` });
  }
  const original = ({ plain: 0, underline: 1, strike: 2, both: 3 } as Record<string, number>)[sample.mode];
  await expect.poll(() => marks(editor)).toEqual(Array(1).fill(original));
  await editor.focus(); await page.keyboard.press('Control+Home');
  for (let i = 0; i < 1; i++) await page.keyboard.press('ArrowRight');
  for (let i = 0; i < 1; i++) await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Control+u'); await page.keyboard.press('Control+Shift+s');
  await expect.poll(() => marks(editor)).toEqual(Array(1).fill(original ^ 3));
  await page.keyboard.press('Control+z'); await page.keyboard.press('Control+z');
  await expect.poll(() => marks(editor)).toEqual(Array(1).fill(original));
  await page.keyboard.press('Control+y'); await page.keyboard.press('Control+y');
  await expect.poll(() => marks(editor)).toEqual(Array(1).fill(original ^ 3));
  if (sample.context === 'stories')
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await waitWordLayout(body);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await capture(page, root, 'edited', hashes);
  const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click();
  await waitWordLayout(body); await capture(page, root, 'reloaded', hashes);
  await page.locator('input[type=file][multiple]').setInputFiles(`${root}/edited.docx`);
  await waitWordLayout(body); await capture(page, root, 'reimported', hashes);
  expect(errors).toEqual([]);
  await fs.writeFile(`${root}/report.json`, JSON.stringify({ name: sample.name, sourceHash: sample.sourceHash,
    hashes, errors, buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-pdf-tab-leader-decorations.spec.ts')),
    scope: 'Browser formatting/history/persistence/actual files; full native geometry comparison is separate.' }, null, 2));
});




