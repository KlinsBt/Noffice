import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

for (const indent of [240, -240]) test(`${indent}twip story indent retains originals through font failure and recovers the same layout`, async ({ page }) => {
  const run = process.env.NOFFICE_LEADER_STORY_RECOVERY_RUN || 'recovery';
  if (!/^recovery(?:-v\d+)?$/.test(run)) throw Error('Invalid leader recovery run');
  const root = `.local/word-tab-leader-stories/${run}/${indent}`;await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-tab-leader-story-arial10.docx');
  const zip = await JSZip.loadAsync(source), header = await zip.file('word/headerLeader.xml')!.async('string');
  expect(header).toContain('w:left="0"');zip.file('word/headerLeader.xml', header.replace('w:left="0"', `w:left="${indent}"`));
  const input = await zip.generateAsync({ type: 'nodebuffer' });await fs.writeFile(root + '/input.docx', input);
  const errors: string[] = [], downloads: string[] = [];page.on('pageerror', (e) => errors.push(e.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');await page.evaluate(() => {
    // Indented leader layout is supported. Exercise an actual font failure
    // explicitly rather than treating a missing browser grant as layout loss.
    (window as unknown as { queryLocalFonts: () => Promise<unknown[]> }).queryLocalFonts = async () => [];
  });await page.locator('input[type=file][multiple]').setInputFiles({ name: 'leader-indent.docx', buffer: input,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });await expect(body).toBeVisible();
  const observe = (d: { suggestedFilename(): string }) => downloads.push(d.suggestedFilename());page.on('download', observe);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('regular font Arial is unavailable for PDF export');
  expect(downloads).toEqual([]);await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();page.off('download', observe);
  await body.focus();await page.keyboard.press('Control+Home');await page.keyboard.insertText('X');
  await expect(body.locator('p').first()).toHaveText('XBody1');
  await page.keyboard.press('Control+z');await expect(body.locator('p').first()).toHaveText('Body1');
  await page.keyboard.press('Control+y');await expect(body.locator('p').first()).toHaveText('XBody1');
  await page.keyboard.press('Control+z');await expect(body.locator('p').first()).toHaveText('Body1');
  await page.getByRole('button', { name: 'Export', exact: true }).click();let pending = page.waitForEvent('download');
  await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
  await (await pending).saveAs(root + '/original.docx');expect(await fs.readFile(root + '/original.docx')).toEqual(input);
  await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'leader-indent', exact: true }).click();await expect(body.locator('p').first()).toHaveText('Body1');
  // Reload restores the real font API. Recover this same indented document
  // before also checking the existing unindented control.
  await page.getByRole('button', { name: 'Export', exact: true }).click();pending = page.waitForEvent('download');
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();await (await pending).saveAs(root + '/indented.pdf');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: 'leader-indent-recovered.docx', buffer: source,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect(page.locator('.word-page-story [data-word-tab-leader-painted=true]')).toHaveCount(20);
  await page.context().grantPermissions(['local-fonts']);await page.getByRole('button', { name: 'Export', exact: true }).click();pending = page.waitForEvent('download');
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();await (await pending).saveAs(root + '/recovered.pdf');
  expect(errors).toEqual([]);
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ inputHash: hash(input), sourceHash: hash(source),
    originalHash: hash(await fs.readFile(root + '/original.docx')), indentedPdfHash: hash(await fs.readFile(root + '/indented.pdf')),
    pdfHash: hash(await fs.readFile(root + '/recovered.pdf')), errors, buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-tab-leader-story-recovery.spec.ts')) }, null, 2));
});
