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
      r.onsuccess = () => {
        const file = r.result.find((file: OfficeFile) => ['Spacing recovery', 'Spacing newer'].includes(file.name));
        if (file) resolve(file); else reject(Error('Missing spacing recovery document'));
      };
      r.onerror = () => reject(r.error);
    }); } finally { db.close(); }
  });
}
for (const failure of ['late-abort', 'other-tab'] as const) test(`Word paragraph spacing survives ${failure} with a usable actual DOCX`, async ({ page, context }) => {
  const root = '.local/word-paragraph-spacing-modes/recovery-browser-v1/' + failure;
  await fs.mkdir(root, { recursive: true });
  const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
  const source = await fs.readFile('tests/fixtures/word-paragraph-spacing-line-before-no-fallback.docx');
  const errors: string[] = []; page.on('pageerror', e => errors.push(e.message));
  await context.grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: 'Spacing recovery.docx', buffer: source,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(body.locator('p')).toHaveCount(5);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const original = await stored(page); let other: Page | undefined;
  if (failure === 'late-abort') await page.evaluate(() => {
    const put = IDBObjectStore.prototype.put;
    IDBObjectStore.prototype.put = function (...args: Parameters<typeof put>) {
      const request = put.apply(this, args);
      if (this.name === 'files') request.addEventListener('success', () => this.transaction.abort(), { once: true });
      return request;
    };
    Object.assign(window, { restoreSpacingSave: () => { IDBObjectStore.prototype.put = put; } });
  });
  else {
    other = await context.newPage(); await other.goto('/');
    await other.getByRole('button', { name: 'Recent files', exact: true }).click();
    await other.getByRole('button', { name: 'Spacing recovery', exact: true }).click();
    await other.getByRole('textbox', { name: 'File name', exact: true }).fill('Spacing newer');
    await expect(other.getByText('Saved on this device', { exact: true })).toBeVisible();
  }
  const edit = async () => {
    await body.locator('p').nth(2).click(); await page.getByRole('button', { name: 'Layout', exact: true }).click();
    const before = page.getByRole('spinbutton', { name: 'Before paragraph', exact: true });
    await before.fill('0.5'); await before.press('Enter');
    await page.getByLabel('After paragraph unit', { exact: true }).selectOption('auto');
  };
  await edit(); await expect(page.getByText('Not saved', { exact: true })).toBeVisible();
  if (failure === 'other-tab') await expect(page.getByRole('alert')).toContainText('another tab');
  else expect(await stored(page)).toEqual(original);
  const outputs: Record<string, string> = {};
  const download = async (name: string, label: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await waiting).saveAs(root + '/' + name); outputs[name] = hash(await fs.readFile(root + '/' + name));
  };
  await download('unsaved.docx', 'DOCX file Editable in Microsoft Word'); await download('unsaved.pdf', 'PDF file');
  if (failure === 'late-abort') {
    await page.evaluate(() => (window as unknown as { restoreSpacingSave: () => void }).restoreSpacingSave());
    await page.keyboard.press('Control+s'); await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    expect((await stored(page)).revision).toBe(original.revision + 1);
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  } else {
    const newer = await stored(page); expect(newer.name).toBe('Spacing newer'); expect(newer.content).toEqual(original.content);
  }
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: failure === 'late-abort' ? 'Spacing recovery' : 'Spacing newer', exact: true }).click();
  if (other) {
    await body.locator('p').nth(2).click(); await page.getByRole('button', { name: 'Layout', exact: true }).click();
    await expect(page.getByRole('spinbutton', { name: 'Before paragraph', exact: true })).toHaveValue('1');
    await page.locator('input[type=file][multiple]').setInputFiles(root + '/unsaved.docx'); await other.close();
  }
  await body.locator('p').nth(2).click(); await page.getByRole('button', { name: 'Layout', exact: true }).click();
  await expect(page.getByRole('spinbutton', { name: 'Before paragraph', exact: true })).toHaveValue('0.5');
  await expect(page.getByLabel('Before paragraph unit', { exact: true })).toHaveValue('lines');
  await expect(page.getByLabel('After paragraph unit', { exact: true })).toHaveValue('auto');
  await download('recovered.pdf', 'PDF file');
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ failure, sourceHash: hash(source), outputs, errors,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-paragraph-spacing-recovery.spec.ts')) }, null, 2));
});
