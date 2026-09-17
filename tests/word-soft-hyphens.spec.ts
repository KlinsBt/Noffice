import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import reference from './fixtures/word-soft-hyphens/reference.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
async function capture(page: Page, root: string, stage: string, hashes: Record<string, string>) {
  for (const [extension, label] of [['docx', 'DOCX file Editable in Microsoft Word'], ['pdf', 'PDF file']]) {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    const path = `${root}/${stage}.${extension}`;
    await (await pending).saveAs(path); hashes[`${stage}.${extension}`] = hash(await fs.readFile(path));
  }
}

for (const sample of reference.rows) test(`Word hyphen identity and lifecycle: ${sample.name}`, async ({ page }) => {
  test.setTimeout(120000);
  const run = process.env.NOFFICE_HYPHEN_RUN || 'browser-v2';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid hyphen evidence folder');
  const root = `.local/word-soft-hyphens/${run}/${sample.name}`;
  await fs.mkdir(root, { recursive: true });
  const hashes: Record<string, string> = {}, errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  expect(hash(await fs.readFile(sample.path))).toBe(sample.sourceHash);
  await page.context().grantPermissions(['local-fonts']);
  await page.setViewportSize({ width: 1500, height: 1400 });
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(sample.path);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await waitWordLayout(editor);
  const model = () => editor.evaluate(el => (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getJSON());
  const text = () => editor.evaluate(el => {
    const editor = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    const paragraphs: string[] = [];
    editor.state.doc.forEach(p => paragraphs.push(p.textBetween(0, p.content.size, '', n => {
      if (n.type.name !== 'wordHyphen') throw Error('Unexpected inline object');
      if (!['optional', 'literal'].includes(n.attrs.kind)) throw Error('Invalid hyphen');
      return n.attrs.kind === 'optional' ? '\u001f' : '\u00ad';
    }) + '\r'));
    return paragraphs.join('');
  });
  expect(await text()).toBe(sample.native.sourceText);
  const sourceModel = await model();
  await fs.writeFile(`${root}/source.html`, await editor.innerHTML());
  await page.screenshot({ path: `${root}/source.png`, fullPage: true });
  await capture(page, root, 'source', hashes); expect(hashes['source.docx']).toBe(sample.sourceHash);
  // Select the actual inline atom, then delete using the browser keyboard.
  // Setting a model range is recorded explicitly; this is not native UI evidence.
  await editor.evaluate(el => {
    const editor = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    let at = -1; editor.state.doc.child(1).forEach((node, offset) => {
      if (at < 0 && node.type.name === 'wordHyphen') at = editor.state.doc.child(0).nodeSize + 1 + offset;
    });
    if (at < 0) throw Error('Missing paragraph2 hyphen');
    editor.commands.setTextSelection({ from: at, to: at + 1 });
  });
  await editor.focus(); await page.keyboard.press('Delete'); await waitWordLayout(editor);
  expect(await text()).toBe(sample.native.editedText); const editedModel = await model();
  await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(sourceModel);
  await page.keyboard.press('Control+y'); await expect.poll(model).toEqual(editedModel);
  await waitWordLayout(editor); await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await capture(page, root, 'edited', hashes);
  const original = await JSZip.loadAsync(await fs.readFile(sample.path));
  const edited = await JSZip.loadAsync(await fs.readFile(`${root}/edited.docx`));
  expect(Object.keys(edited.files).sort()).toEqual(Object.keys(original.files).sort());
  for (const [path, entry] of Object.entries(original.files)) {
    if (entry.dir || ['word/document.xml', 'word/settings.xml'].includes(path)) continue;
    expect(await edited.file(path)!.async('uint8array')).toEqual(await entry.async('uint8array'));
  }
  const title = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: title, exact: true }).click(); await waitWordLayout(editor);
  expect(await text()).toBe(sample.native.editedText); await capture(page, root, 'reloaded', hashes);
  await page.locator('input[type=file][multiple]').setInputFiles(`${root}/edited.docx`); await waitWordLayout(editor);
  expect(await text()).toBe(sample.native.editedText); await capture(page, root, 'reimported', hashes);
  expect(hashes['reimported.docx']).toBe(hashes['edited.docx']);
  const beforeFailure = await model(), downloads: string[] = [];
  const download = (item: { suggestedFilename(): string }) => downloads.push(item.suggestedFilename());
  page.on('download', download);
  await page.evaluate(() => {
    const target = window as any; target.hyphenFonts = target.queryLocalFonts; target.queryLocalFonts = async () => [];
  });
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('font Arial is unavailable');
  expect(downloads).toEqual([]); expect(await model()).toEqual(beforeFailure);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click(); page.off('download', download);
  await editor.focus(); await page.keyboard.press('Control+End'); await page.keyboard.type('X'); const typed = await model();
  expect(typed).not.toEqual(beforeFailure);
  await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(beforeFailure);
  await page.keyboard.press('Control+y'); await expect.poll(model).toEqual(typed);
  await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(beforeFailure);
  // Only copies returned to this test are changed. Installed font files stay intact.
  await page.evaluate(async () => {
    const target = window as any, native = await target.hyphenFonts.call(window);
    target.hyphenFontRecords = [];
    target.queryLocalFonts = async () => native.map((font: any) => ({ family: font.family, style: font.style,
      blob: async () => {
        const bytes = await (await font.blob()).arrayBuffer(), view = new DataView(bytes);
        let found = false;
        for (let i = 0; i < view.getUint16(4); i++) {
          const entry = 12 + i * 16;
          if (view.getUint32(entry) === 0x4f532f32) {
            view.setUint16(view.getUint32(entry + 8) + 8, 0x100); found = true;
          }
        }
        if (!found) throw Error('Missing OS/2 font test table');
        const digest = await crypto.subtle.digest('SHA-256', bytes);
        target.hyphenFontRecords.push({ family: font.family, style: font.style, size: bytes.byteLength,
          hash: [...new Uint8Array(digest)].map(byte => byte.toString(16).padStart(2, '0')).join('') });
        return new Blob([bytes]);
      } }));
  });
  await waitWordLayout(editor); await capture(page, root, 'recovered', hashes);
  expect(hashes['recovered.docx']).toBe(hashes['edited.docx']); expect(errors).toEqual([]);
  const fontRecords = await page.evaluate(() => (window as any).hyphenFontRecords);
  expect(new Set(fontRecords.map((font: { family: string; style: string }) => `${font.family}|${font.style}`)).size).toBe(2);
  await fs.writeFile(`${root}/report.json`, JSON.stringify({ name: sample.name, sourceHash: sample.sourceHash,
    sourceModel, editedModel, hashes, fontRecords, errors, buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-soft-hyphens.spec.ts')),
    scope: 'Optional versus literal soft-hyphen identity; import, selected browser keyboard deletion, undo/redo, reload, actual DOCX/PDF downloads, reimport, missing-font failure and continued editing/recovery. Native output, glyph/paint and native UI acceptance are separate.' }, null, 2));
});
