import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

test('marked tab leader font failure recovers without losing mark history or reload', async ({ page }) => {
  test.setTimeout(90000);
  const run = process.env.NOFFICE_PDF_TAB_RECOVERY_RUN || 'recovery-browser-v1';
  if (!/^recovery-browser-v\d+$/.test(run)) throw Error('Invalid tab recovery run');
  const root = `.local/word-pdf-tab-decorations/${run}`; await fs.mkdir(root, { recursive: true });
  const zip = await JSZip.loadAsync(await fs.readFile('tests/fixtures/word-pdf-tab-decorations/body-both.docx'));
  const source = await zip.file('word/document.xml')!.async('string');
  expect(source).toContain('w:val="left"');
  zip.file('word/document.xml', source.replace('w:val="left"', 'w:val="left" w:leader="dot"'));
  const bytes = await zip.generateAsync({ type: 'nodebuffer' }); await fs.writeFile(`${root}/source.docx`, bytes);
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(`${root}/source.docx`);
  const body = page.getByRole('textbox', { name: 'Document text', exact: true }); await waitWordLayout(body);
  const original = await body.innerText(), downloads: string[] = [];
  page.on('download', download => downloads.push(download.suggestedFilename()));
  await page.evaluate(() => {
    const target = window as any; target.tabOriginalFonts = target.queryLocalFonts;
    target.queryLocalFonts = async () => (await target.tabOriginalFonts.call(window))
      .filter((font: any) => font.family !== 'Calibri');
  });
  const reject = async () => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    await page.getByRole('button', { name: 'PDF file', exact: true }).click();
    await expect(page.getByRole('alert')).toContainText('regular font Calibri is unavailable');
    await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
    await expect.poll(() => body.innerText()).toBe(original);
  };
  await reject(); expect(downloads).toEqual([]);
  await page.evaluate(() => { const target = window as any; target.queryLocalFonts = target.tabOriginalFonts; });
  await body.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Control+u'); await page.keyboard.press('Control+Shift+s');
  const capture = async (stage: string) => {
    for (const [extension, label] of [['docx', 'DOCX file Editable in Microsoft Word'], ['pdf', 'PDF file']]) {
      await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: label, exact: true }).click(); await (await pending).saveAs(`${root}/${stage}.${extension}`);
    }
  };
  await waitWordLayout(body); await capture('recovered');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.evaluate(() => { const target = window as any;
    target.queryLocalFonts = async () => (await target.tabOriginalFonts.call(window))
      .filter((font: any) => font.family !== 'Calibri'); });
  await reject();
  expect(downloads).toHaveLength(2);
  await page.evaluate(() => { const target = window as any; target.queryLocalFonts = target.tabOriginalFonts; });
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click(); await waitWordLayout(body); await capture('reloaded');
  await expect.poll(() => body.innerText()).toBe(original);
  const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
  const hashes: Record<string, string> = {};
  for (const name of ['source.docx', 'recovered.docx', 'recovered.pdf', 'reloaded.docx', 'reloaded.pdf'])
    hashes[name] = hash(await fs.readFile(`${root}/${name}`));
  await fs.writeFile(`${root}/report.json`, JSON.stringify({ hashes, downloads, buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-pdf-tab-decoration-failure.spec.ts')),
    scope: 'Missing-font rejection with no failure download or content mutation; remove marks, undo/reject, restore fonts, redo/reload and actual recovered exports. Marked leaders use the same measured spans.' }, null, 2));
});
