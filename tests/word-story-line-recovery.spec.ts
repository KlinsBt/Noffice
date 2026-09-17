import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';
import { wordStoryBuildHash } from './word-story-artifacts';

async function stored(page: Page) {
  return page.evaluate(async () => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const r = indexedDB.open('noffice-workspace', 1); r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error);
    });
    try { return await new Promise<OfficeFile>((resolve, reject) => {
      const r = db.transaction('files').objectStore('files').getAll();
      r.onsuccess = () => { const file = r.result.find((f: OfficeFile) => f.name.startsWith('Line recovery')); if (file) resolve(file); else reject(Error('Missing recovery document')); };
      r.onerror = () => reject(r.error);
    }); } finally { db.close(); }
  });
}
for (const failure of ['late-abort', 'other-tab'] as const) test(`Word footer line spacing preserves unsaved edits after ${failure}`, async ({ page, context }) => {
  const run = process.env.NOFFICE_STORY_LINE_RECOVERY_RUN || 'recovery-browser-v1';
  expect(run).toMatch(/^recovery-browser-v\d+$/);
  const root = `.local/word-story-line-authoring/${run}/${failure}`; await fs.mkdir(root, { recursive: true });
  const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
  const input = await fs.readFile('tests/fixtures/word-story-spacing-grid-footer-before-3-0.docx');
  const errors: string[] = []; page.on('pageerror', e => errors.push(e.message)); const outputs: Record<string, string> = {};
  await context.grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: 'Line recovery.docx', buffer: input,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect(page.locator('.section-page')).toHaveCount(2); await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const original = await stored(page); let other: Page | undefined;
  if (failure === 'late-abort') await page.evaluate(() => {
    const put = IDBObjectStore.prototype.put;
    IDBObjectStore.prototype.put = function (...args: Parameters<typeof put>) {
      const r = put.apply(this, args); if (this.name === 'files') r.addEventListener('success', () => this.transaction.abort(), { once: true }); return r;
    };
    Object.assign(window, { restoreLineSave: () => { IDBObjectStore.prototype.put = put; } });
  });
  else {
    other = await context.newPage(); await other.goto('/'); await other.getByRole('button', { name: 'Recent files', exact: true }).click();
    await other.getByRole('button', { name: 'Line recovery', exact: true }).click();
    await other.getByRole('textbox', { name: 'File name', exact: true }).fill('Line recovery newer');
    await expect(other.getByText('Saved on this device', { exact: true })).toBeVisible();
  }
  const open = async () => {
    await page.getByRole('button', { name: 'Insert', exact: true }).click(); await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: /^Section 1 \u2014 Default footer/ }).click();
    await page.getByRole('textbox', { name: 'Header or footer text', exact: true }).locator('p').nth(2).click();
    return page.getByRole('dialog').getByRole('combobox', { name: 'Line spacing', exact: true });
  };
  await (await open()).selectOption('2'); await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await expect(page.getByText('Not saved', { exact: true })).toBeVisible();
  if (failure === 'late-abort') expect(await stored(page)).toEqual(original);
  else { await expect(page.getByRole('alert')).toContainText('another tab'); const newer = await stored(page); expect(newer.name).toBe('Line recovery newer'); expect(newer.content).toEqual(original.content); }
  const download = async (file: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click(); await (await waiting).saveAs(root + '/' + file); outputs[file] = hash(await fs.readFile(root + '/' + file));
  };
  await download('unsaved.docx', 'DOCX file Editable in Microsoft Word'); await download('unsaved.pdf');
  if (failure === 'late-abort') {
    await page.evaluate(() => (window as unknown as { restoreLineSave: () => void }).restoreLineSave()); await page.keyboard.press('Control+s');
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible(); expect((await stored(page)).revision).toBe(original.revision + 1);
  }
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: failure === 'late-abort' ? 'Line recovery' : 'Line recovery newer', exact: true }).click();
  if (other) { await expect(await open()).toHaveValue('40pt'); await page.getByRole('dialog').getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.locator('input[type=file][multiple]').setInputFiles(root + '/unsaved.docx'); await other.close(); }
  await expect(await open()).toHaveValue('2'); await page.getByRole('dialog').getByRole('button', { name: 'Cancel', exact: true }).click();
  await download('recovered.pdf'); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ failure, sourceHash: hash(input), outputs, errors,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-story-line-recovery.spec.ts')) }, null, 2));
});
