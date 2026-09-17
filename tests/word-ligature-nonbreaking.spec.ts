import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

test('Word PDF preserves NBSP and ligatures with full-font embedding and failure recovery', async ({ page }) => {
  const run = process.env.NOFFICE_NBSP_FULL_FONT_RUN || 'browser-v2';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid font evidence folder');
  const root = `.local/word-ligature-nonbreaking/${run}`;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-ligature-nonbreaking.docx');
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const reference = JSON.parse(await fs.readFile('tests/fixtures/native-word-ligature-nonbreaking.json', 'utf8'));
  expect(hash(source)).toBe(reference.sourceHash);
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: 'NBSP full font ligature.docx',
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect(page.locator('.section-page')).toHaveCount(2);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const subset = page.waitForEvent('download');
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await (await subset).saveAs(root + '/subset.pdf');
  // Change only a returned in-memory font copy, never the installed font.
  const fonts = await page.evaluate(async () => {
    const target = window as any, query = target.queryLocalFonts;
    target.fullFontPermission = 0x100;
    const native = await query.call(window), records: { hash: string; size: number }[] = [];
    target.queryLocalFonts = async () => native.map((font: any) => ({ family: font.family, style: font.style,
      blob: async () => {
        const bytes = await (await font.blob()).arrayBuffer();
        if (font.family === 'Calibri' && font.style === 'Regular') {
          const data = new DataView(bytes); let found = false;
          for (let i = 0; i < data.getUint16(4); i++) {
            const offset = 12 + i * 16;
            if (data.getUint32(offset) === 0x4f532f32) {
              data.setUint16(data.getUint32(offset + 8) + 8, target.fullFontPermission); found = true;
            }
          }
          if (!found) throw Error('Missing OS/2 test table');
          const digest = await crypto.subtle.digest('SHA-256', bytes);
          records.push({ hash: [...new Uint8Array(digest)].map(b => b.toString(16).padStart(2, '0')).join(''), size: bytes.byteLength });
        }
        return new Blob([bytes]);
      } }));
    target.fullFontRecords = records;
    return native.filter((font: any) => font.family === 'Calibri' && font.style === 'Regular').length;
  });
  expect(fonts).toBe(1);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download', { timeout: 10000 });
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await (await pending).saveAs(root + '/full.pdf');
  const rejectedDownloads: string[] = [];
  page.on('download', download => rejectedDownloads.push(download.suggestedFilename()));
  await page.evaluate(() => { (window as any).fullFontPermission = 2; });
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('does not permit embedding');
  expect(rejectedDownloads).toEqual([]);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  await page.evaluate(() => { (window as any).fullFontPermission = 0x100; });
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: /^Section 1 \u2014 Default footer/ }).click();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.press('ArrowRight');
  await page.keyboard.insertText('X'); await expect(editor.locator('p').first()).toHaveText('tXi\u00a0ti ti\u00a0\u00a0ti\tB');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText('ti\u00a0ti ti\u00a0\u00a0ti\tB');
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText('tXi\u00a0ti ti\u00a0\u00a0ti\tB');
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  const footer = page.locator('.word-page-story[data-word-story*="footer"] p').first();
  await expect(footer).toHaveText('tXi\u00a0ti ti\u00a0\u00a0ti\tB');
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(footer).toHaveText('ti\u00a0ti ti\u00a0\u00a0ti\tB');
  await page.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(footer).toHaveText('tXi\u00a0ti ti\u00a0\u00a0ti\tB');
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(footer).toHaveText('ti\u00a0ti ti\u00a0\u00a0ti\tB');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const restored = page.waitForEvent('download');
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await (await restored).saveAs(root + '/restored.pdf');
  const fontRecords = await page.evaluate(() => (window as any).fullFontRecords);
  expect(fontRecords.length).toBeGreaterThan(0);
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'NBSP full font ligature', exact: true }).click();
  await expect(page.locator('.word-page-story[data-word-story*="footer"] p').first()).toHaveText('ti\u00a0ti ti\u00a0\u00a0ti\tB');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const original = page.waitForEvent('download');
  await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
  await (await original).saveAs(root + '/original.docx');
  expect(await fs.readFile(root + '/original.docx')).toEqual(source);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(source),
    subsetHash: hash(await fs.readFile(root + '/subset.pdf')),
    pdfHash: hash(await fs.readFile(root + '/full.pdf')), restoredHash: hash(await fs.readFile(root + '/restored.pdf')),
    fontRecords, buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-ligature-nonbreaking.spec.ts')) }, null, 2));
});
