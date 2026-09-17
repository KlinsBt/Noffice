import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';
import { contentFingerprint } from '../src/office-preservation';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
const run = process.env.NOFFICE_SECTION_MIGRATION_RUN || 'migration-browser-v1';
if (!/^migration-browser-v\d+$/.test(run)) throw Error('Invalid migration run');
async function stored(page: Page, name: string) {
  return page.evaluate(async name => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('noffice-workspace', 1);
      request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error);
    });
    try { return await new Promise<OfficeFile>((resolve, reject) => {
      const request = db.transaction('files').objectStore('files').getAll();
      request.onsuccess = () => resolve(request.result.find((file: OfficeFile) => file.name === name));
      request.onerror = () => reject(request.error);
    }); } finally { db.close(); }
  }, name);
}

for (const action of ['unchanged', 'edited', 'explicit', 'broken-original']) test(`Word legacy final section ${action} retains edits, revisions and recoverable originals`, async ({ page }) => {
  const root = `.local/word-implicit-final-sections/${run}/${action}`;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-implicit-final-3-sections-omitted.docx');
  const name = 'Legacy final ' + action;
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(body).toBeVisible();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const file = await stored(page, name);
  if (file.content.kind !== 'word') throw Error('Expected Word');
  // The old importer used the preceding A4 landscape section for these controls.
  // Unit coverage independently reconstructs this baseline with the old reader.
  const legacy = { ...file.content, orientation: 'landscape' as const };
  delete legacy.sectionDefaultsVersion;
  const fingerprint = await contentFingerprint(legacy);
  if (action === 'edited' || action === 'explicit') legacy.html = legacy.html.replace('first paragraph.', 'edited paragraph.');
  if (action === 'explicit') legacy.pageOverrides = { orientation: true };
  await page.evaluate(async ({ name, legacy, fingerprint, broken }) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('noffice-workspace', 1);
      request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error);
    });
    try { await new Promise<void>((resolve, reject) => {
      const tx = db.transaction('files', 'readwrite'), store = tx.objectStore('files'), request = store.getAll();
      request.onsuccess = () => {
        const file = request.result.find((file: OfficeFile) => file.name === name) as OfficeFile;
        file.content = legacy; file.original!.contentFingerprint = fingerprint;
        if (broken) file.original!.data = new Uint8Array(new TextEncoder().encode('preserved invalid original')).buffer;
        store.put(file);
      };
      tx.oncomplete = () => resolve(); tx.onerror = () => reject(tx.error); tx.onabort = () => reject(tx.error);
    }); } finally { db.close(); }
  }, { name, legacy, fingerprint, broken: action === 'broken-original' });
  const reopen = async () => {
    await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
  };
  await reopen();
  const outputs: Record<string, string> = {};
  const download = async (fileName: string, label: string | RegExp, recovery = false) => {
    if (!recovery) await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: typeof label === 'string' }).click();
    await (await pending).saveAs(root + '/' + fileName);
    outputs[fileName] = hash(await fs.readFile(root + '/' + fileName));
  };
  if (action === 'broken-original') {
    await expect(page.getByRole('region', { name: 'Document recovery' }).getByRole('alert')).toBeVisible();
    await expect(body).toHaveCount(0);
    await download('recovery.noffice', 'Export Noffice backup', true);
    await download('original.docx', 'Download original DOCX', true);
    const backup = JSON.parse(await fs.readFile(root + '/recovery.noffice', 'utf8'));
    expect(backup.content).toEqual(legacy);
    expect(Buffer.from(backup.original.base64, 'base64').toString()).toBe('preserved invalid original');
    expect(await fs.readFile(root + '/original.docx', 'utf8')).toBe('preserved invalid original');
  } else {
    await expect(body.locator('p').first()).toHaveText(`Section 1 ${action === 'unchanged' ? 'first' : 'edited'} paragraph.`);
    await download('migrated.noffice', /^Noffice backup/);
    const migrated = JSON.parse(await fs.readFile(root + '/migrated.noffice', 'utf8'));
    expect(migrated.content.sectionDefaultsVersion).toBe(1);
    expect(migrated.content.orientation).toBe(action === 'explicit' ? 'landscape' : 'portrait');
    expect(migrated.content.html).toBe(legacy.html);
    expect(migrated.content.pageOverrides).toEqual(legacy.pageOverrides);
    expect(migrated.original.base64).toBe(source.toString('base64'));
    await download('migrated.docx', 'DOCX file Editable in Microsoft Word');
    if (action === 'unchanged') expect(outputs['migrated.docx']).toBe(hash(source));
    await reopen();
    await expect(body.locator('p').first()).toHaveText(`Section 1 ${action === 'unchanged' ? 'first' : 'edited'} paragraph.`);
  }
  const after = await stored(page, name);
  expect(after.revision).toBe(file.revision); expect(after.content).toEqual(legacy);
  expect(after.original!.contentFingerprint).toBe(fingerprint);
  await fs.writeFile(root + '/report.json', JSON.stringify({ action, outputs, revision: after.revision,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-section-defaults-migration.spec.ts')) }, null, 2));
});
