import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';
import { contentFingerprint } from '../src/office-preservation';
import { wordStoryBuildHash } from './word-story-artifacts';

for (const kind of ['uniform-edited', 'mixed-unchanged', 'mixed-edited'] as const)
  test(`legacy font features ${kind}: persisted recovery and retained backup`, async ({ page }) => {
    const mixed = kind.startsWith('mixed'), edited = kind !== 'mixed-unchanged';
    const name = `Feature recovery ${kind}`;
    const source = await fs.readFile(`tests/fixtures/word-ligature-boundary-${mixed ? 'disabled-second' : 'single'}.docx`);
    const run = process.env.NOFFICE_FONT_FEATURE_RECOVERY_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
    const root = `.local/word-ligatures/recovery/${run}/${kind}`;
    await fs.mkdir(root, { recursive: true });
    const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx',
      buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const storyPath = await page.locator('.word-page-story[data-word-story*="footer"]').first().getAttribute('data-word-story');
    expect(storyPath).toBeTruthy();
    const stored = (strip: boolean) => page.evaluate(({ strip, name, edited, storyPath }) => new Promise<OfficeFile>((resolve, reject) => {
      const request = indexedDB.open('noffice-workspace');
      request.onerror = () => reject(request.error);
      request.onsuccess = () => {
        const db = request.result, tx = db.transaction('files', strip ? 'readwrite' : 'readonly');
        const store = tx.objectStore('files'), all = store.getAll(); let result: OfficeFile;
        all.onsuccess = () => {
          result = all.result.find((file: OfficeFile) => file.name === name);
          if (!result || result.content.kind !== 'word' || !result.content.stories) { tx.abort(); return; }
          if (strip) {
            const clean = (html: string, change = false) => {
              const dom = new DOMParser().parseFromString(html, 'text/html');
              for (const element of dom.querySelectorAll<HTMLElement>('[data-word-font-features],[data-word-paragraph-font-features]')) {
                element.removeAttribute('data-word-font-features');
                element.removeAttribute('data-word-paragraph-font-features');
                element.style.removeProperty('font-feature-settings');
              }
              if (change) dom.querySelector('p')!.append(' saved edit');
              return dom.body.innerHTML;
            };
            delete result.content.fontFeaturesVersion;
            result.content.html = clean(result.content.html);
            for (const part of result.content.stories.parts) {
              part.html = clean(part.html, edited && part.path === storyPath);
            }
            store.put(result);
          }
        };
        tx.oncomplete = () => { db.close(); resolve(result); };
        tx.onerror = tx.onabort = () => { db.close(); reject(tx.error); };
      };
    }), { strip, name, edited, storyPath });
    const stripped = await stored(true);
    if (!edited) {
      // A genuine unchanged older snapshot fingerprints its older HTML. Keeping
      // the modern fingerprint would seed an edited file instead.
      const fingerprint = await contentFingerprint(stripped.content);
      await page.evaluate(({ name, fingerprint }) => new Promise<void>((resolve, reject) => {
        const request = indexedDB.open('noffice-workspace');
        request.onerror = () => reject(request.error);
        request.onsuccess = () => {
          const db = request.result, tx = db.transaction('files', 'readwrite');
          const store = tx.objectStore('files'), all = store.getAll();
          all.onsuccess = () => {
            const file = all.result.find((value: OfficeFile) => value.name === name) as OfficeFile;
            file.original!.contentFingerprint = fingerprint; store.put(file);
          };
          tx.oncomplete = () => { db.close(); resolve(); };
          tx.onerror = tx.onabort = () => { db.close(); reject(tx.error); };
        };
      }), { name, fingerprint });
    }
    const before = await stored(false);
    await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    const failure = kind === 'mixed-edited';
    if (failure) {
      await expect(page.getByRole('region', { name: 'Document recovery' })
        .getByRole('alert')).toHaveText(/edited text with missing mixed font-feature metadata/);
      await expect(page.getByRole('textbox', { name: 'Document text', exact: true })).toHaveCount(0);
    } else {
      const footer = page.locator('.word-page-story[data-word-story*="footer"] p').first();
      await expect(footer).toHaveText('ti\tB' + (edited ? ' saved edit' : ''));
      const values = await footer.locator('[data-word-font-features]').evaluateAll(elements =>
        [...new Set(elements.map(element => element.getAttribute('data-word-font-features')))]);
      expect(values.sort()).toEqual(mixed ? ['0', '1'] : ['1']);
    }
    expect(await stored(false)).toEqual(before);
    const download = async (label: string, filename: string) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: label, exact: true }).click();
      await (await pending).saveAs(root + '/' + filename);
      return fs.readFile(root + '/' + filename);
    };
    const backupBytes = await download('Noffice backup Full editable model and retained original · .noffice', 'recovered.noffice');
    const backup = JSON.parse(backupBytes.toString());
    expect(Buffer.from(backup.original.base64, 'base64')).toEqual(source);
    expect(backup.content.fontFeaturesVersion).toBe(failure ? undefined : 1);
    if (failure) expect(backup.content).toEqual(before.content);
    else {
      const docx = await download('DOCX file Editable in Microsoft Word', 'recovered.docx');
      if (!edited) expect(docx).toEqual(source);
      await page.goto('/');
      await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '-export.docx',
        buffer: docx, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      await expect(page.locator('.word-page-story[data-word-story*="footer"] p').first())
        .toHaveText('ti\tB' + (edited ? ' saved edit' : ''));
    }
    await fs.writeFile(root + '/browser-report.json', JSON.stringify({ kind, passed: true,
      sourceHash: hash(source), backupHash: hash(backupBytes), unchangedStorageThroughHydration: true,
      buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-font-feature-recovery.spec.ts')) }, null, 2));
  });
