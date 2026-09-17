import { test, expect, type Locator, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/word-pdf-decorations/reference.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout } from './word-pagination-helpers';
import JSZip from 'jszip';

const run = process.env.NOFFICE_PDF_DECORATIONS_RUN || 'browser-v1';
if (!/^browser-v\d+$/.test(run)) throw Error('Invalid PDF decoration run');
const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
async function marks(editor: Locator) {
  return editor.locator('p').first().evaluate(p => {
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), result: number[] = [];
    while (walker.nextNode()) {
      const text = walker.currentNode as Text;
      result.push(...Array.from(text.data, () => (text.parentElement?.closest('u') ? 1 : 0)
        | (text.parentElement?.closest('s,strike,del') ? 2 : 0)));
    }
    return result.slice(4, 23);
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
  const root = `.local/word-pdf-decorations/${run}/${sample.name}`;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile(sample.fixture); expect(hash(source)).toBe(sample.sourceHash);
  const hashes: Record<string, string> = {}, errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(sample.fixture);
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  await waitWordLayout(body); await capture(page, root, 'source', hashes);
  expect(hashes['source.docx']).toBe(sample.sourceHash);
  let editor = body;
  if (sample.context === 'stories') {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: /^Section 1 .*Default header/ }).click();
    editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  }
  const original = ({ plain: 0, underline: 1, strike: 2, both: 3 } as Record<string, number>)[sample.mode];
  await expect.poll(() => marks(editor)).toEqual(Array(19).fill(original));
  await editor.focus(); await page.keyboard.press('Control+Home');
  for (let i = 0; i < 4; i++) await page.keyboard.press('ArrowRight');
  for (let i = 0; i < 19; i++) await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Control+u'); await page.keyboard.press('Control+Shift+s');
  await expect.poll(() => marks(editor)).toEqual(Array(19).fill(original ^ 3));
  await page.keyboard.press('Control+z'); await page.keyboard.press('Control+z');
  await expect.poll(() => marks(editor)).toEqual(Array(19).fill(original));
  await page.keyboard.press('Control+y'); await page.keyboard.press('Control+y');
  await expect.poll(() => marks(editor)).toEqual(Array(19).fill(original ^ 3));
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
    hashes, errors, buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-pdf-decorations.spec.ts')),
    scope: 'Browser formatting/history/persistence/actual files; full native geometry comparison is separate.' }, null, 2));
});

test('Word PDF held-out64 size and font-face source controls', async ({ page }) => {
  const root = `.local/word-pdf-decorations/${run}/heldout`; await fs.mkdir(root, { recursive: true });
  const source = 'tests/fixtures/word-pdf-decorations/heldout.docx';
  expect(hash(await fs.readFile(source))).toBe(reference.heldoutHash);
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(source);
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  await waitWordLayout(body); await expect(body.locator('p')).toHaveCount(64);
  const hashes: Record<string, string> = {}; await capture(page, root, 'source', hashes);
  await fs.writeFile(`${root}/report.json`, JSON.stringify({ hashes, buildHash: await wordStoryBuildHash(),
    sourceHash: reference.heldoutHash, testHash: hash(await fs.readFile('tests/word-pdf-decorations.spec.ts')) }, null, 2));
});

for (const [name, property] of [['double', '<w:u w:val="double"/>'], ['wave', '<w:u w:val="wave"/>'],
  ['color', '<w:u w:val="single" w:color="FF0000"/>'], ['double-strike', '<w:dstrike/>']])
  test(`Word PDF rejects retained ${name} source loss without losing edits or originals`, async ({ page }) => {
    const root = `.local/word-pdf-decorations/${run}/failure-${name}`; await fs.mkdir(root, { recursive: true });
    const zip = await JSZip.loadAsync(await fs.readFile(reference.rows[0].fixture));
    zip.file('word/document.xml', (await zip.file('word/document.xml')!.async('string'))
      .replaceAll('<w:rPr>', '<w:rPr>' + property));
    const bytes = await zip.generateAsync({ type: 'nodebuffer' });
    await fs.writeFile(`${root}/source.docx`, bytes);
    await page.context().grantPermissions(['local-fonts']); await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/source.docx`);
    const body = page.getByRole('textbox', { name: 'Document text', exact: true }); await waitWordLayout(body);
    const originalText = await body.innerText(); const downloads: string[] = [];
    page.on('download', download => downloads.push(download.suggestedFilename()));
    const reject = async () => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      await page.getByRole('button', { name: 'PDF file', exact: true }).click();
      await expect(page.getByRole('alert')).toContainText('complex underline or double-strikethrough');
      await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
    };
    await reject(); expect(downloads).toEqual([]);
    await body.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.type('Kept ');
    await expect(body).toContainText('Kept P01'); await reject(); expect(downloads).toEqual([]);
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect.poll(() => body.innerText()).toBe(originalText);
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
    const path = `${root}/unchanged.docx`; await (await pending).saveAs(path);
    expect(await fs.readFile(path)).toEqual(bytes);
    await fs.writeFile(`${root}/report.json`, JSON.stringify({ sourceHash: hash(bytes), originalHash: hash(await fs.readFile(path)),
      buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-pdf-decorations.spec.ts')),
      scope: 'PDF rejection, no failure download, retained editing/undo and byte-identical original DOCX recovery.' }, null, 2));
  });
