import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';
import { contentFingerprint } from '../src/office-preservation';
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout } from './word-pagination-helpers';
import reference from './fixtures/word-soft-hyphens/reference.json' with { type: 'json' };

const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
async function stored(page: Page, name: string) {
  return page.evaluate(async name => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const r = indexedDB.open('noffice-workspace', 1); r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error);
    });
    try { return await new Promise<OfficeFile>((resolve, reject) => {
      const r = db.transaction('files').objectStore('files').getAll();
      r.onsuccess = () => resolve(r.result.find((file: OfficeFile) => file.name === name)); r.onerror = () => reject(r.error);
    }); } finally { db.close(); }
  }, name);
}

for (const encoding of ['element', 'literal']) for (const action of ['unchanged', 'prefix', 'ambiguous', 'broken-original'])
  test(`Legacy Word hyphen recovery: ${encoding} / ${action}`, async ({ page }) => {
    test.setTimeout(90000);
    const sample = reference.rows.find(row => row.name === `regular-160-${encoding}`)!;
    const name = `Legacy hyphen ${encoding} ${action}`, run = process.env.NOFFICE_HYPHEN_MIGRATION_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Invalid migration evidence run');
    const root = `.local/word-soft-hyphen-migration/${run}/${encoding}-${action}`; await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(sample.path); expect(hash(source)).toBe(sample.sourceHash);
    await page.context().grantPermissions(['local-fonts']); await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true }); await waitWordLayout(editor);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible(); const file = await stored(page, name);
    if (file.content.kind !== 'word') throw Error('Expected Word content');
    const legacy = structuredClone(file.content); delete legacy.hyphenVersion;
    // Reconstruct the pre-hyphen model, which exposed both OOXML encodings as
    // raw U+00AD. Independent importer unit tests also exercise that old path.
    legacy.html = await page.evaluate(html => {
      const doc = new DOMParser().parseFromString(html, 'text/html');
      for (const marker of doc.querySelectorAll('[data-word-hyphen]')) marker.replaceWith(doc.createTextNode('\u00ad'));
      return doc.body.innerHTML;
    }, legacy.html);
    const fingerprint = await contentFingerprint(legacy);
    if (action === 'prefix') legacy.html = legacy.html.replace('antidisestab', 'Edited antidisestab');
    if (action === 'ambiguous') legacy.html = legacy.html.replace('micro\u00adscope', 'new\u00adtext');
    await page.evaluate(async ({ name, legacy, fingerprint, broken }) => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const r = indexedDB.open('noffice-workspace', 1); r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error);
      });
      try { await new Promise<void>((resolve, reject) => {
        const tx = db.transaction('files', 'readwrite'), store = tx.objectStore('files'), r = store.getAll();
        r.onsuccess = () => {
          const file = r.result.find((file: OfficeFile) => file.name === name) as OfficeFile;
          file.content = legacy; file.original!.contentFingerprint = fingerprint;
          if (broken) file.original!.data = Uint8Array.from(new TextEncoder().encode('preserved invalid hyphen source')).buffer;
          store.put(file);
        };
        tx.oncomplete = () => resolve(); tx.onerror = () => reject(tx.error); tx.onabort = () => reject(tx.error);
      }); } finally { db.close(); }
    }, { name, legacy, fingerprint, broken: action === 'broken-original' });
    const reopen = async () => {
      await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
      await page.getByRole('button', { name, exact: true }).click();
    };
    await reopen(); const hashes: Record<string, string> = {};
    const download = async (fileName: string, label: string | RegExp, recovery = false) => {
      if (!recovery) await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download'); await page.getByRole('button', { name: label, exact: typeof label === 'string' }).click();
      const path = root + '/' + fileName; await (await pending).saveAs(path); hashes[fileName] = hash(await fs.readFile(path));
    };
    if (['ambiguous', 'broken-original'].includes(action)) {
      const recovery = page.getByRole('region', { name: 'Document recovery' }); await expect(recovery.getByRole('alert')).toBeVisible();
      if (action === 'ambiguous') await expect(recovery.getByRole('alert')).toContainText('ambiguous edited hyphen');
      await expect(editor).toHaveCount(0);
      await download('recovery.noffice', 'Export Noffice backup', true); await download('original.docx', 'Download original DOCX', true);
      const backup = JSON.parse(await fs.readFile(root + '/recovery.noffice', 'utf8')); expect(backup.content).toEqual(legacy);
      const expected = action === 'broken-original' ? Buffer.from('preserved invalid hyphen source') : source;
      expect(Buffer.from(backup.original.base64, 'base64')).toEqual(expected); expect(await fs.readFile(root + '/original.docx')).toEqual(expected);
    } else {
      await waitWordLayout(editor); await download('migrated.noffice', /^Noffice backup/);
      const migrated = JSON.parse(await fs.readFile(root + '/migrated.noffice', 'utf8'));
      expect(migrated.content.hyphenVersion).toBe(1); expect(migrated.original.base64).toBe(source.toString('base64'));
      const markers = await editor.locator('[data-word-hyphen]').evaluateAll(nodes => nodes.map(n => n.getAttribute('data-word-hyphen')));
      expect(markers).toHaveLength(22); expect(new Set(markers)).toEqual(new Set([encoding === 'element' ? 'optional' : 'literal']));
      if (action === 'prefix') await expect(editor.locator('p').nth(1)).toContainText('Edited antidisestab');
      await download('migrated.docx', 'DOCX file Editable in Microsoft Word'); await download('migrated.pdf', 'PDF file');
      if (action === 'unchanged') expect(hashes['migrated.docx']).toBe(sample.sourceHash);
      await reopen(); await waitWordLayout(editor); expect(await editor.locator('[data-word-hyphen]').count()).toBe(22);
    }
    const after = await stored(page, name); expect(after.revision).toBe(file.revision); expect(after.content).toEqual(legacy);
    expect(after.original!.contentFingerprint).toBe(fingerprint);
    await fs.writeFile(root + '/report.json', JSON.stringify({ name, action, encoding, hashes, legacy,
      revision: after.revision, sourceHash: sample.sourceHash, buildHash: await wordStoryBuildHash(),
      testHash: hash(await fs.readFile('tests/word-hyphen-migration.spec.ts')),
      scope: 'Reconstructed old raw-U+00AD model: unchanged/prefix migration, ambiguous edit and corrupt-original recovery; actual backup/DOCX/PDF files and unchanged stored revisions. Native output comparison is separate.' }, null, 2));
  });
