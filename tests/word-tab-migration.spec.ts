import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';
import { wordStoryBuildHash } from './word-story-artifacts';
const reference = JSON.parse(readFileSync('tests/fixtures/native-word-tab-legacy.json', 'utf8')) as typeof import('./fixtures/native-word-tab-legacy.json');
const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');

const samples = [
  ...reference.rows.map((sample) => ({ ...sample, missingSource: false })),
  { ...reference.rows.find((sample) => sample.name === 'stories-edited')!, name: 'stories-missing-source', missingSource: true },
];
for (const sample of samples)
  test(`${sample.missingSource ? 'damaged' : 'genuine'} pre-tab ${sample.name} backup preserves edits and storage through migration or recovery`, async ({ page }) => {
    const input = await fs.readFile('tests/fixtures/' + sample.file); expect(hash(input)).toBe(sample.backupHash);
    const backup = JSON.parse(input.toString('utf8')); expect(backup.content.tabStopsVersion).toBeUndefined();
    // Deliberately damage only copied-story provenance in a real edited backup.
    // Recovery must retain the exact damaged payload so the user can repair it.
    if (sample.missingSource) backup.content.stories.parts[0].copiedFrom = 'word/missing-tab-source.xml';
    const name = 'Legacy tabs ' + sample.name, root = '.local/word-tab-legacy/browser/' + sample.name;
    await fs.mkdir(root, { recursive: true });
    const errors: string[] = []; page.on('pageerror', (e) => errors.push(e.message));
    await page.setViewportSize({ width: 1500, height: 1200 }); await page.goto('/');
    await page.evaluate(async ({ backup, name }) => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const r = indexedDB.open('noffice-workspace', 1);
        r.onupgradeneeded = () => {
          r.result.createObjectStore('files', { keyPath: 'id' });
          r.result.createObjectStore('versions', { keyPath: 'key' }).createIndex('fileId', 'fileId');
        };
        r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error);
      });
      try {
        await new Promise<void>((resolve, reject) => {
          const tx = db.transaction('files', 'readwrite');
          tx.objectStore('files').put({ id: name, name, kind: 'word', content: backup.content,
            createdAt: 1, updatedAt: 1, revision: 7, favorite: false, trashed: false, warnings: backup.warnings,
            ...(backup.original ? { original: { name: backup.original.name, contentFingerprint: backup.original.contentFingerprint,
              data: Uint8Array.from(atob(backup.original.base64), (c) => c.charCodeAt(0)).buffer } } : {}),
          });
          tx.oncomplete = () => resolve(); tx.onabort = () => reject(tx.error); tx.onerror = () => reject(tx.error);
        });
      } finally { db.close(); }
    }, { backup, name });
    await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    if (sample.missingSource) {
      const recovery = page.getByRole('region', { name: 'Document recovery' });
      await expect(recovery.getByRole('alert')).toContainText('no source for its tab stops');
      await expect(page.getByRole('textbox', { name: 'Document text', exact: true })).toHaveCount(0);
      await expect(page.getByText('Opening document...', { exact: true })).toHaveCount(0);
      let waiting = page.waitForEvent('download');
      await recovery.getByRole('button', { name: 'Export Noffice backup', exact: true }).click();
      await (await waiting).saveAs(root + '/recovered.noffice');
      const recovered = JSON.parse(await fs.readFile(root + '/recovered.noffice', 'utf8'));
      expect(recovered.content).toEqual(backup.content);
      expect(recovered.original.base64).toBe(backup.original.base64);
      waiting = page.waitForEvent('download');
      await recovery.getByRole('button', { name: 'Download original DOCX', exact: true }).click();
      await (await waiting).saveAs(root + '/original.docx');
      expect(hash(await fs.readFile(root + '/original.docx'))).toBe(sample.sourceHash);
      const stored = await page.evaluate(async (name) => {
        const db = await new Promise<IDBDatabase>((resolve, reject) => { const r = indexedDB.open('noffice-workspace', 1);r.onsuccess = () => resolve(r.result);r.onerror = () => reject(r.error); });
        try { return await new Promise<OfficeFile>((resolve, reject) => { const r = db.transaction('files').objectStore('files').get(name);r.onsuccess = () => resolve(r.result);r.onerror = () => reject(r.error); }); }
        finally { db.close(); }
      }, name);
      expect(stored.revision).toBe(7);expect(stored.content).toEqual(backup.content);
      expect(stored.original?.contentFingerprint).toBe(backup.original.contentFingerprint);
      expect(errors).toEqual([]);
      await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(input),
        damage: 'copiedFrom points to word/missing-tab-source.xml', recoveredHash: hash(await fs.readFile(root + '/recovered.noffice')),
        originalHash: sample.sourceHash, rejected: true, revision: stored.revision, errors, buildHash: await wordStoryBuildHash() }, null, 2));
      return;
    }
    const body = page.getByRole('textbox', { name: 'Document text', exact: true }); await expect(body).toBeVisible();
    if (sample.kind === 'body') {
      await expect(body.locator('p').first()).toHaveText((sample.edited ? 'Saved ' : '') + 'A\tB\tC');
      await expect(body.locator('p').first().locator('[data-word-tab-measured=true]')).toHaveCount(2);
    } else {
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
      await expect(page.getByRole('textbox', { name: 'Header or footer text', exact: true })).toHaveText(
        sample.kind === 'fresh' ? 'Saved \tCenter\tRight' : (sample.edited ? 'Saved ' : '') + 'Left\tCenter\tRight');
      await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    }
    const outputs: Record<string, string> = {};
    const download = async (name: string, label: string | RegExp) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click(); const waiting = page.waitForEvent('download');
      await page.getByRole('button', { name: label, exact: typeof label === 'string' }).click();
      await (await waiting).saveAs(root + '/' + name); outputs[name] = hash(await fs.readFile(root + '/' + name));
    };
    await download('recovered.noffice', /^Noffice backup/);
    const recovered = JSON.parse(await fs.readFile(root + '/recovered.noffice', 'utf8'));
    expect(recovered.content.tabStopsVersion).toBe(1);
    expect(recovered.original?.base64).toBe(backup.original?.base64);
    if (sample.kind === 'fresh') expect(recovered.original).toBeUndefined();
    await download('recovered.docx', 'DOCX file Editable in Microsoft Word');
    if (!sample.edited) expect(outputs['recovered.docx']).toBe(sample.sourceHash);
    const stored = await page.evaluate(async (name) => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => { const r = indexedDB.open('noffice-workspace', 1);r.onsuccess = () => resolve(r.result);r.onerror = () => reject(r.error); });
      try { return await new Promise<OfficeFile>((resolve, reject) => { const r = db.transaction('files').objectStore('files').get(name);r.onsuccess = () => resolve(r.result);r.onerror = () => reject(r.error); }); }
      finally { db.close(); }
    }, name);
    expect(stored.revision).toBe(7); expect(stored.content).toEqual(backup.content);
    expect(stored.original?.contentFingerprint).toBe(backup.original?.contentFingerprint);
    expect(errors).toEqual([]);
    await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(input), outputs,
      revision: stored.revision, errors, buildHash: await wordStoryBuildHash() }, null, 2));
  });
