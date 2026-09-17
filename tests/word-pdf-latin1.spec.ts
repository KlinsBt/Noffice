import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import reference from './fixtures/word-pdf-latin1/reference.json' with { type: 'json' };
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

for (const sample of reference.rows) test(`Latin-1 ${sample.name}: keyboard edit, history, files and full-font recovery`, async ({ page }) => {
  test.setTimeout(120000);
  const run = process.env.NOFFICE_PDF_LATIN1_RUN || 'browser-v4';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid Latin-1 evidence folder');
  const root = `.local/word-latin1-pdf/${run}/${sample.name}`;
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
  const text = () => editor.locator('p').allTextContents();
  const original = sample.controls.map(control => control.text), edited = [...original];
  edited[0] = edited[0].replace('\u00a1', '\u00bf');
  expect(await text()).toEqual(original); await capture(page, root, 'source', hashes);
  expect(hashes['source.docx']).toBe(sample.sourceHash);
  await editor.focus(); await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowRight'); await page.keyboard.press('Shift+ArrowRight');
  expect(await editor.evaluate(el => {
    const editor = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    return editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to, '');
  })).toBe('\u00a1');
  await page.keyboard.insertText('\u00bf'); await waitWordLayout(editor);
  await expect.poll(text).toEqual(edited);
  const initialSession = (await model()).attrs!.wordInitialEditSession;
  expect(initialSession).toMatch(/^[a-f\d]{32}$/);
  await page.keyboard.press('Control+z'); await expect.poll(text).toEqual(original);
  expect((await model()).attrs!.wordInitialEditSession).toBeNull();
  await page.keyboard.press('Control+y'); await expect.poll(text).toEqual(edited);
  expect((await model()).attrs!.wordInitialEditSession).toBe(initialSession);
  await waitWordLayout(editor);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await capture(page, root, 'edited', hashes);
  const originalZip = await JSZip.loadAsync(await fs.readFile(sample.path));
  const editedZip = await JSZip.loadAsync(await fs.readFile(`${root}/edited.docx`));
  expect(Object.keys(editedZip.files).sort()).toEqual(Object.keys(originalZip.files).sort());
  for (const [path, entry] of Object.entries(originalZip.files)) {
    if (entry.dir || ['word/document.xml', 'word/settings.xml'].includes(path)) continue;
    expect(await editedZip.file(path)!.async('uint8array')).toEqual(await entry.async('uint8array'));
  }
  const xml = await editedZip.file('word/document.xml')!.async('string');
  const originalSettings = await originalZip.file('word/settings.xml')!.async('string');
  const savedSettings = await editedZip.file('word/settings.xml')!.async('string');
  expect(await page.evaluate(({ original, saved }) => {
    const parser = new DOMParser(), serializer = new XMLSerializer();
    const before = parser.parseFromString(original, 'application/xml');
    const after = parser.parseFromString(saved, 'application/xml');
    const registers = [...after.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'rsids')];
    if (registers.length !== 1) return false;
    registers[0].remove();
    return serializer.serializeToString(before) === serializer.serializeToString(after);
  }, { original: originalSettings, saved: savedSettings })).toBe(true);
  expect(await page.evaluate((xml) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    return [...document.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't')]
      .filter(text => text.getAttributeNS('http://www.w3.org/XML/1998/namespace', 'space') === 'preserve' &&
        !/^[ \t\r\n]|[ \t\r\n]$/.test(text.textContent || '')).length;
  }, xml)).toBe(0);
  const title = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: title, exact: true }).click(); await waitWordLayout(editor);
  expect((await model()).attrs!.wordInitialEditSession).toBe(initialSession);
  expect(await text()).toEqual(edited); await capture(page, root, 'reloaded', hashes);
  await editor.focus(); await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowRight'); await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.insertText('\u00a1'); await waitWordLayout(editor);
  expect(await text()).toEqual(original);
  expect((await model()).attrs!.wordInitialEditSession).toBe(initialSession);
  await capture(page, root, 'second-edit', hashes);
  await editor.focus(); await expect(editor).toBeFocused();
  await page.keyboard.press('Control+z'); await expect.poll(text).toEqual(edited);
  await page.keyboard.press('Control+y'); await expect.poll(text).toEqual(original);
  await page.keyboard.press('Control+z'); await expect.poll(text).toEqual(edited);
  expect((await model()).attrs!.wordInitialEditSession).toBe(initialSession);
  await page.locator('input[type=file][multiple]').setInputFiles(`${root}/edited.docx`);
  await waitWordLayout(editor); expect(await text()).toEqual(edited);
  await capture(page, root, 'reimported', hashes);

  const beforeFailure = await model(), downloads: string[] = [];
  const rejectedDownload = (download: { suggestedFilename(): string }) => downloads.push(download.suggestedFilename());
  page.on('download', rejectedDownload);
  await page.evaluate(() => {
    const target = window as any; target.latin1Fonts = target.queryLocalFonts; target.queryLocalFonts = async () => [];
  });
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('font Arial is unavailable');
  expect(downloads).toEqual([]); expect(await model()).toEqual(beforeFailure);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  page.off('download', rejectedDownload);
  await editor.focus(); await page.keyboard.press('Control+End'); await page.keyboard.type('X');
  const typed = await model(); expect(typed).not.toEqual(beforeFailure);
  await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(beforeFailure);
  await page.keyboard.press('Control+y'); await expect.poll(model).toEqual(typed);
  await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(beforeFailure);
  // Only copies returned to this test are changed. Installed font files stay intact.
  await page.evaluate(async () => {
    const target = window as any, native = await target.latin1Fonts.call(window);
    target.latin1FontRecords = [];
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
        target.latin1FontRecords.push({ family: font.family, style: font.style, size: bytes.byteLength,
          hash: [...new Uint8Array(digest)].map(byte => byte.toString(16).padStart(2, '0')).join('') });
        return new Blob([bytes]);
      } }));
  });
  await waitWordLayout(editor); await capture(page, root, 'recovered', hashes);
  expect(hashes['recovered.docx']).toBe(hashes['edited.docx']);
  const fontRecords = await page.evaluate(() => (window as any).latin1FontRecords);
  expect(new Set(fontRecords.map((font: { family: string; style: string }) => `${font.family}|${font.style}`)).size).toBe(2);
  expect(errors).toEqual([]);
  await fs.writeFile(`${root}/report.json`, JSON.stringify({ name: sample.name, sourceHash: sample.sourceHash, hashes,
    fontRecords, errors, buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-pdf-latin1.spec.ts')),
    scope: 'All94 printable Latin-1 characters in both native fonts/four faces. Keyboard punctuation replacement, initial/later save sessions, undo/redo, reload/reimport and DOCX/PDF downloads. Missing-font failure preserves the model/history; recovered PDF exercises full embedding. Native file/glyph/font-stream comparisons are separate.' }, null, 2));
});
