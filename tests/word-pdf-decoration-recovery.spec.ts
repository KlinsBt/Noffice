import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

test('decorated PDF recovers from a missing styled face and invalid font metrics, including offline export', async ({ page, context }) => {
  const run = process.env.NOFFICE_PDF_DECORATION_RECOVERY_RUN || 'recovery-browser-v1';
  if (!/^recovery-browser-v\d+$/.test(run)) throw Error('Invalid decoration recovery run');
  const root = `.local/word-pdf-decorations/${run}`; await fs.mkdir(root, { recursive: true });
  const source = 'tests/fixtures/word-pdf-decorations/body-both.docx';
  await context.grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(source);
  const body = page.getByRole('textbox', { name: 'Document text', exact: true }); await expect(body).toBeVisible();
  const original = await body.innerText(); const downloads: string[] = [];
  page.on('download', download => downloads.push(download.suggestedFilename()));
  await page.evaluate(() => {
    const target = window as any; target.decorationOriginalFonts = target.queryLocalFonts;
    target.queryLocalFonts = async () => (await target.decorationOriginalFonts.call(window))
      .filter((font: any) => !(font.family === 'Arial' && /^bold[ -]?italic$/i.test(font.style)));
  });
  const start = async () => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  };
  await start(); await expect(page.getByRole('alert')).toContainText('boldItalic font Arial is unavailable');
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  expect(downloads).toEqual([]);
  await page.evaluate(() => {
    const target = window as any;
    target.queryLocalFonts = async () => (await target.decorationOriginalFonts.call(window)).map((font: any) => ({
      family: font.family, style: font.style, blob: async () => {
        const blob = await font.blob();
        if (font.family !== 'Arial' || font.style !== 'Regular') return blob;
        const bytes = await blob.arrayBuffer(), view = new DataView(bytes); let changed = false;
        for (let i = 0; i < view.getUint16(4); i++) {
          const entry = 12 + i * 16;
          if (view.getUint32(entry) === 0x4f532f32) { view.setInt16(view.getUint32(entry + 8) + 26, 0); changed = true; }
        }
        if (!changed) throw Error('Missing control OS/2 table');
        return new Blob([bytes]);
      },
    }));
  });
  await start(); await expect(page.getByRole('alert')).toContainText('invalid underline or strikethrough metrics');
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  expect(downloads).toEqual([]); await expect.poll(() => body.innerText()).toBe(original);
  await page.evaluate(async () => { const target = window as any; target.queryLocalFonts = target.decorationOriginalFonts;
    await navigator.serviceWorker.ready; });
  await context.setOffline(true);
  const pending = page.waitForEvent('download'); await start(); await (await pending).saveAs(root + '/recovered.pdf');
  await expect.poll(() => body.innerText()).toBe(original);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  await fs.writeFile(root + '/report.json', JSON.stringify({ buildHash: await wordStoryBuildHash(),
    sourceHash: hash(await fs.readFile(source)), pdfHash: hash(await fs.readFile(root + '/recovered.pdf')),
    testHash: hash(await fs.readFile('tests/word-pdf-decoration-recovery.spec.ts')), downloads,
    scope: 'Actual worker rejection for missing boldItalic face and zero strike metrics; no error downloads, unchanged content, restored font API and successful offline eight-face PDF export.' }, null, 2));
});
