import { storyKeyboardTests } from './word-story-keyboard-workflow';
import { test, expect } from '@playwright/test';
import type { OfficeFile } from '../src/model';
storyKeyboardTests(
  ['missing-normal', 'missing-character', 'missing-both', 'missing-styles-part'],
  'word-story-base',
  '.local/word-story-create-base-styles/browser',
);

test('legacy unavailable base-style templates recover before editing without background saves', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles('tests/fixtures/word-story-base-missing-both.docx');
  const name = 'word-story-base-missing-both';
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const before = await page.evaluate(async (name) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const r = indexedDB.open('noffice-workspace', 1);
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
    try {
      return await new Promise<{ revision: number; fingerprint: string }>((resolve, reject) => {
        const tx = db.transaction('files', 'readwrite'),
          store = tx.objectStore('files'),
          r = store.getAll();
        let result: { revision: number; fingerprint: string };
        r.onsuccess = () => {
          const file = (r.result as OfficeFile[]).find((f) => f.name === name)!;
          if (file.content.kind !== 'word') throw Error('Expected Word');
          file.content.stories!.templateVersion = 2;
          file.content.stories!.emptyTemplates = { header: null, footer: null };
          file.content.html = file.content.html.replace('B1-01', 'Saved edit');
          result = { revision: file.revision, fingerprint: file.original!.contentFingerprint! };
          store.put(file);
        };
        tx.oncomplete = () => resolve(result);
        tx.onabort = () => reject(tx.error);
        tx.onerror = () => reject(tx.error);
      });
    } finally {
      db.close();
    }
  }, name);
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click();
  await expect(page.getByRole('textbox', { name: 'Document text', exact: true })).toContainText(
    'Saved edit',
  );
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  const choice = page.getByRole('button', { name: /^Section 1 — First-page header/ });
  await expect(choice).toBeEnabled();
  await choice.click();
  await expect(
    page.getByRole('textbox', { name: 'Header or footer text', exact: true }),
  ).toHaveText('');
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  const stored = await page.evaluate(async (name) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const r = indexedDB.open('noffice-workspace', 1);
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
    try {
      return await new Promise<OfficeFile>((resolve, reject) => {
        const r = db.transaction('files').objectStore('files').getAll();
        r.onsuccess = () => resolve(r.result.find((f: OfficeFile) => f.name === name));
        r.onerror = () => reject(r.error);
      });
    } finally {
      db.close();
    }
  }, name);
  expect(stored.revision).toBe(before.revision);
  expect(stored.original?.contentFingerprint).toBe(before.fingerprint);
  if (stored.content.kind !== 'word') throw Error('Expected Word');
  expect(stored.content.stories?.templateVersion).toBe(2);
  expect(stored.content.html).toContain('Saved edit');
});
