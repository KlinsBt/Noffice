import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';
import { wordStoryBuildHash } from './word-story-artifacts';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-style-defaults-legacy.json', 'utf8'),
) as typeof import('./fixtures/native-word-style-defaults-legacy.json');
const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');

for (const sample of reference.rows)
  test(`real v3 ${sample.name} snapshot preserves edits and original bytes through migration or recovery`, async ({
    page,
  }) => {
    const input = await fs.readFile('tests/fixtures/' + sample.file);
    expect(hash(input)).toBe(sample.backupHash);
    const backup = JSON.parse(input.toString('utf8'));
    expect(backup.content.styleDefaultsVersion).toBeUndefined();
    expect(backup.content.stories.templateVersion).toBe(3);
    const rejected = sample.edited && !sample.name.startsWith('unaffected');
    const name = 'Legacy ' + sample.name;
    const root = '.local/word-style-defaults/migration/' + sample.name;
    await fs.mkdir(root, { recursive: true });
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.goto('/');
    // Seed the exact previous build's saved content. An ambiguous migration
    // must leave this transaction/revision untouched and allow a backup.
    await page.evaluate(
      async ({ backup, name }) => {
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          const r = indexedDB.open('noffice-workspace', 1);
          r.onupgradeneeded = () => {
            r.result.createObjectStore('files', { keyPath: 'id' });
            r.result
              .createObjectStore('versions', { keyPath: 'key' })
              .createIndex('fileId', 'fileId');
          };
          r.onsuccess = () => resolve(r.result);
          r.onerror = () => reject(r.error);
        });
        try {
          await new Promise<void>((resolve, reject) => {
            const tx = db.transaction('files', 'readwrite');
            tx.objectStore('files').put({
              id: name,
              name,
              kind: 'word',
              content: backup.content,
              createdAt: 1,
              updatedAt: 1,
              revision: 7,
              favorite: false,
              trashed: false,
              warnings: backup.warnings,
              original: {
                name: backup.original.name,
                contentFingerprint: backup.original.contentFingerprint,
                data: Uint8Array.from(atob(backup.original.base64), (c) => c.charCodeAt(0)).buffer,
              },
            });
            tx.oncomplete = () => resolve();
            tx.onabort = () => reject(tx.error);
            tx.onerror = () => reject(tx.error);
          });
        } finally {
          db.close();
        }
      },
      { backup, name },
    );
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    if (sample.edited && !rejected) await expect(editor).toContainText('Saved legacy edit');
    if (rejected)
      await expect(
        page.getByRole('region', { name: 'Document recovery' }).getByRole('alert'),
      ).toBeVisible();
    else {
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      const choice = page.getByRole('button', { name: /^Section 1 \u2014 First-page header/ });
      await expect(choice).toBeEnabled();
      await choice.click();
      await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    }
    let waiting = page.waitForEvent('download');
    if (rejected) {
      await expect(editor).toHaveCount(0);
      await expect(page.getByText('Opening document...', { exact: true })).toHaveCount(0);
      await page.getByRole('button', { name: 'Export Noffice backup', exact: true }).click();
    } else {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      await page.getByRole('button', { name: /^Noffice backup/ }).click();
    }
    await (await waiting).saveAs(root + '/recovered.noffice');
    const recovered = JSON.parse(await fs.readFile(root + '/recovered.noffice', 'utf8'));
    expect(recovered.original.base64).toBe(backup.original.base64);
    if (rejected) expect(recovered.content).toEqual(backup.content);
    else {
      expect(recovered.content.styleDefaultsVersion).toBe(1);
      expect(recovered.content.stories.templateVersion).toBe(4);
      if (sample.edited) expect(recovered.content.html).toBe(backup.content.html);
    }
    if (!sample.edited) {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      waiting = page.waitForEvent('download');
      await page
        .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
        .click();
      await (await waiting).saveAs(root + '/unchanged.docx');
      expect(hash(await fs.readFile(root + '/unchanged.docx'))).toBe(sample.sourceHash);
    }
    if (rejected) {
      waiting = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Download original DOCX', exact: true }).click();
      await (await waiting).saveAs(root + '/original.docx');
      expect(hash(await fs.readFile(root + '/original.docx'))).toBe(sample.sourceHash);
    }
    const stored = await page.evaluate(async (name) => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const r = indexedDB.open('noffice-workspace', 1);
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
      try {
        return await new Promise<OfficeFile>((resolve, reject) => {
          const r = db.transaction('files').objectStore('files').get(name);
          r.onsuccess = () => resolve(r.result);
          r.onerror = () => reject(r.error);
        });
      } finally {
        db.close();
      }
    }, name);
    expect(stored.revision).toBe(7);
    expect(stored.content).toEqual(backup.content);
    expect(stored.original?.contentFingerprint).toBe(backup.original.contentFingerprint);
    expect(errors).toEqual([]);
    await fs.writeFile(
      root + '/browser-report.json',
      JSON.stringify(
        {
          sourceHash: hash(input),
          buildHash: await wordStoryBuildHash(),
          recoveredHash: hash(await fs.readFile(root + '/recovered.noffice')),
          rejected,
          revision: stored.revision,
          errors,
        },
        null,
        2,
      ),
    );
  });
