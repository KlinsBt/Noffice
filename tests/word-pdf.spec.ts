import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { PDFDocument } from 'pdf-lib';

const root = '.local/word-pdf-recovery';
async function startExport(page: Page) {
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
}
async function savePdf(page: Page, name: string) {
  const pending = page.waitForEvent('download');
  await startExport(page);
  const download = await pending;
  const bytes = await fs.readFile((await download.path())!);
  const pdf = await PDFDocument.load(bytes);
  expect(pdf.getPageCount()).toBe(1);
  expect(pdf.getPage(0).getWidth()).toBeCloseTo(595.3, 2);
  expect(pdf.getPage(0).getHeight()).toBeCloseTo(841.9, 2);
  await fs.writeFile(`${root}/${name}.pdf`, bytes);
}

test('Word PDF exports empty and short pages offline with bundled fonts and formatting recovery', async ({
  page,
  context,
}) => {
  await fs.mkdir(root, { recursive: true });
  const external: string[] = [];
  page.on('request', (request) => {
    if (
      new URL(request.url()).origin !== 'http://127.0.0.1:4183' &&
      !request.url().startsWith('data:')
    )
      external.push(request.url());
  });
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Word', exact: true }).click();
  await page.evaluate(async () => {
    await navigator.serviceWorker.ready;
  });
  await expect.poll(() => page.evaluate(() => !!navigator.serviceWorker.controller)).toBe(true);
  await context.setOffline(true);
  await savePdf(page, 'empty');
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.fill('Short PDF page');
  await page.keyboard.press('Control+a');
  await page.keyboard.press('Control+b');
  await startExport(page);
  await expect(page.getByRole('alert')).toContainText('does not yet support this font style');
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  for (const family of ['Inter', 'Source Serif 4', 'JetBrains Mono']) {
    await editor.focus();
    await page.keyboard.press('Control+a');
    await page.getByRole('combobox', { name: 'Font family', exact: true }).selectOption(family);
    await savePdf(page, family.replaceAll(' ', '-'));
  }
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Untitled document', exact: true }).click();
  await expect(editor).toHaveText('Short PDF page');
  await savePdf(page, 'reloaded');
  expect(external).toEqual([]);
});

test('Word PDF worker rejects malformed and restricted fonts, times out and discards stale output with keyboard recovery', async ({
  page,
  context,
}) => {
  await fs.mkdir(root, { recursive: true });
  await context.grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Word', exact: true }).click();
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.fill('Recoverable PDF');
  await page.keyboard.press('Control+a');
  await page.getByRole('combobox', { name: 'Font family', exact: true }).selectOption('Arial');
  // Inject a corrupt local-font response at the browser API boundary; the real
  // worker must reject it without corrupting the document or emitting a file.
  await page.evaluate(() => {
    const target = window as any;
    target.originalPdfFonts = target.queryLocalFonts;
    target.queryLocalFonts = async () => [
      { family: 'Arial', style: 'Regular', blob: async () => new Blob([new Uint8Array(12)]) },
    ];
  });
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await startExport(page);
  await expect(page.getByRole('alert')).toContainText('TrueType or OpenType');
  await expect(editor).toHaveText('Recoverable PDF');
  expect(downloads).toEqual([]);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  // Modify only the returned in-memory font copy. The OS font is untouched.
  await page.evaluate(() => {
    const target = window as any;
    target.queryLocalFonts = async () => {
      const fonts = await target.originalPdfFonts.call(window);
      const font = fonts.find((face: any) => face.family === 'Arial' && face.style === 'Regular');
      const bytes = await (await font.blob()).arrayBuffer();
      const data = new DataView(bytes);
      let found = false;
      for (let i = 0; i < data.getUint16(4); i++) {
        const offset = 12 + i * 16;
        if (data.getUint32(offset) === 0x4f532f32) {
          data.setUint16(data.getUint32(offset + 8) + 8, 2);
          found = true;
        }
      }
      if (!found) throw Error('Missing OS/2 test table');
      return [{ family: 'Arial', style: 'Regular', blob: async () => new Blob([bytes]) }];
    };
  });
  await startExport(page);
  await expect(page.getByRole('alert')).toContainText('does not permit embedding');
  expect(downloads).toEqual([]);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  await page.evaluate(() => {
    const target = window as any;
    target.queryLocalFonts = async () => {
      const fonts = await target.originalPdfFonts.call(window);
      return fonts.map((font: any) => ({
        family: font.family,
        style: font.style,
        blob: () =>
          new Promise((resolve) => {
            target.releasePdfFont = async () => resolve(await font.blob());
          }),
      }));
    };
  });
  await startExport(page);
  await expect
    .poll(() => page.evaluate(() => typeof (window as any).releasePdfFont))
    .toBe('function');
  // Keyboard input can arrive while an asynchronous export is pending.
  await editor.click();
  await page.keyboard.press('Control+End');
  await page.keyboard.insertText(' edited');
  await expect(editor).toHaveText('Recoverable PDF edited');
  await page.evaluate(() => (window as any).releasePdfFont());
  await expect(page.getByRole('alert')).toContainText('Document changed before export');
  expect(downloads).toEqual([]);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  await page.evaluate(() => {
    (window as any).queryLocalFonts = (window as any).originalPdfFonts;
  });
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor).toHaveText('Recoverable PDF');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(editor).toHaveText('Recoverable PDF edited');
  // A nonresponding worker exercises the real 15-second deadline and cleanup.
  await page.evaluate(() => {
    const target = window as any;
    target.originalPdfWorker = target.Worker;
    target.pdfWorkerTerminations = 0;
    target.Worker = class extends EventTarget {
      postMessage() {}
      terminate() {
        target.pdfWorkerTerminations++;
      }
    };
  });
  await startExport(page);
  await expect(page.getByRole('alert')).toContainText('PDF export took too long', {
    timeout: 17000,
  });
  expect(await page.evaluate(() => (window as any).pdfWorkerTerminations)).toBe(1);
  expect(downloads).toEqual([]);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  await page.evaluate(() => {
    (window as any).Worker = (window as any).originalPdfWorker;
  });
  await savePdf(page, 'recovered');
});
