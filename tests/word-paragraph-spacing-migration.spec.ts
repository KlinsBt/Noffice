import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';
import { wordStoryBuildHash } from './word-story-artifacts';
import reference from './fixtures/native-word-paragraph-spacing-legacy.json' with { type: 'json' };

const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
async function stored(page: Page, name: string) {
  return page.evaluate(async name => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const r = indexedDB.open('noffice-workspace', 1); r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error);
    });
    try { return await new Promise<OfficeFile>((resolve, reject) => {
      const r = db.transaction('files').objectStore('files').get(name); r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error);
    }); } finally { db.close(); }
  }, name);
}
for (const sample of reference.rows) test(`Word legacy paragraph spacing ${sample.name}: migration preserves saved edits and original recovery`, async ({ page }) => {
  const input = await fs.readFile('tests/fixtures/' + sample.fixture); expect(hash(input)).toBe(sample.backupHash);
  const backup = JSON.parse(input.toString('utf8')), name = 'Spacing legacy ' + sample.name;
  expect(backup.content.paragraphSpacingVersion).toBeUndefined();
  const root = '.local/word-paragraph-spacing-modes/legacy-browser-v1/' + sample.name;
  await fs.mkdir(root, { recursive: true });
  const errors: string[] = []; page.on('pageerror', e => errors.push(e.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.evaluate(async ({ backup, name }) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const r = indexedDB.open('noffice-workspace', 1);
      r.onupgradeneeded = () => {
        r.result.createObjectStore('files', { keyPath: 'id' });
        r.result.createObjectStore('versions', { keyPath: 'key' }).createIndex('fileId', 'fileId');
      };
      r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error);
    });
    try { await new Promise<void>((resolve, reject) => {
      const tx = db.transaction('files', 'readwrite');
      tx.objectStore('files').put({ id: name, name, kind: 'word', content: backup.content,
        createdAt: 1, updatedAt: 1, revision: 7, favorite: false, trashed: false, warnings: backup.warnings,
        original: { name: backup.original.name, contentFingerprint: backup.original.contentFingerprint,
          data: Uint8Array.from(atob(backup.original.base64), c => c.charCodeAt(0)).buffer } });
      tx.oncomplete = () => resolve(); tx.onabort = () => reject(tx.error); tx.onerror = () => reject(tx.error);
    }); } finally { db.close(); }
  }, { backup, name });
  const open = async () => {
    await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
  };
  await open();
  const outputs: Record<string, string> = {};
  const download = async (file: string, label: string | RegExp, recovery = false) => {
    if (!recovery) await page.getByRole('button', { name: 'Export', exact: true }).click();
    const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: typeof label === 'string' }).click();
    await (await waiting).saveAs(root + '/' + file); outputs[file] = hash(await fs.readFile(root + '/' + file));
  };
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  if (sample.edit === 'split') {
    await expect(page.getByRole('region', { name: 'Document recovery' }).getByRole('alert')).toContainText('ambiguous paragraph spacing');
    await expect(body).toHaveCount(0);
    await download('recovered.noffice', 'Export Noffice backup', true);
    await download('original.docx', 'Download original DOCX', true);
    expect(outputs['original.docx']).toBe(sample.sourceHash);
    const recovered = JSON.parse(await fs.readFile(root + '/recovered.noffice', 'utf8'));
    expect(recovered.content).toEqual(backup.content); expect(recovered.original.base64).toBe(backup.original.base64);
  } else {
    await expect(body.locator(':scope > p')).toHaveCount(5);
    await expect(body.locator('p').first()).toHaveText(sample.edit === 'points' ? 'Edited' : 'A');
    await body.locator('p').nth(2).click();
    await page.getByRole('button', { name: 'Layout', exact: true }).click();
    if (sample.name.startsWith('line-')) {
      await expect(page.getByLabel('Before paragraph unit', { exact: true })).toHaveValue(sample.edit === 'points' ? 'points' : 'lines');
      await expect(page.getByRole('spinbutton', { name: 'Before paragraph', exact: true })).toHaveValue(sample.edit === 'points' ? '8' : '1');
    } else if (sample.name.startsWith('auto-')) {
      await expect(page.getByLabel('Before paragraph unit', { exact: true })).toHaveValue('auto');
      await expect(page.getByLabel('After paragraph unit', { exact: true })).toHaveValue('auto');
    } else await expect(page.getByRole('checkbox', { name: "Don't add space between paragraphs of the same style", exact: true })).toBeChecked();
    await download('migrated.noffice', /^Noffice backup/);
    const migrated = JSON.parse(await fs.readFile(root + '/migrated.noffice', 'utf8'));
    expect(migrated.content.paragraphSpacingVersion).toBe(1);
    expect(migrated.original.base64).toBe(backup.original.base64);
    await download('migrated.docx', 'DOCX file Editable in Microsoft Word');
    if (sample.edit === 'none') expect(outputs['migrated.docx']).toBe(sample.sourceHash);
    await open(); await expect(body.locator('p').first()).toHaveText(sample.edit === 'points' ? 'Edited' : 'A');
  }
  const record = await stored(page, name);
  expect(record.revision).toBe(7); expect(record.content).toEqual(backup.content);
  expect(record.original?.contentFingerprint).toBe(backup.original.contentFingerprint);
  if (sample.edit !== 'split') {
    await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/migrated.docx');
    await expect(body.locator('p').first()).toHaveText(sample.edit === 'points' ? 'Edited' : 'A');
    await body.locator('p').nth(2).click(); await page.getByRole('button', { name: 'Layout', exact: true }).click();
    if (sample.edit === 'points') {
      await expect(page.getByLabel('Before paragraph unit', { exact: true })).toHaveValue('points');
      await expect(page.getByRole('spinbutton', { name: 'Before paragraph', exact: true })).toHaveValue('8');
    }
    await download('reimported.pdf', 'PDF file');
  }
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sample, outputs, revision: record.revision, errors,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-paragraph-spacing-migration.spec.ts')) }, null, 2));
});
