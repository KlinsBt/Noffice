import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';
import { contentFingerprint } from '../src/office-preservation';
import reference from './fixtures/native-word-script-scopes.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows.filter(row => row.scope === 'whole-paragraph'))
  for (const mode of ['unchanged', 'edited', 'split'] as const)
    test(`Word paragraph script recovery ${row.name} ${mode}: preserve saved edits and original bytes`, async ({ page }) => {
      const run = process.env.NOFFICE_SCRIPT_RECOVERY_RUN || 'browser-v1';
      if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
      const root = `.local/word-script-scopes/recovery/${run}/${row.name}-${mode}`;
      await fs.mkdir(root, { recursive: true });
      const source = await fs.readFile(`tests/fixtures/word-script-scopes/${row.name}-formatted.docx`);
      const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
      expect(hash(source)).toBe(row.stages.find(stage => stage.stage === 'formatted')!.docxHash);
      const name = `Script recovery ${row.name} ${mode}`, body = row.kind === 'Body';
      const kind = row.kind === 'Headers' ? 'header' : 'footer', errors: string[] = [];
      page.on('pageerror', error => errors.push(error.message));
      await page.context().grantPermissions(['local-fonts']); await page.goto('/');
      await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      const openStory = async () => {
        if (body) return;
        await page.getByRole('button', { name: 'Insert', exact: true }).click();
        await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
        await page.getByRole('button', { name: new RegExp(`^Section 1 — Default ${kind}`) }).click();
      };
      const editor = page.getByRole('textbox', { name: body ? 'Document text' : 'Header or footer text', exact: true });
      const snapshot = () => editor.evaluate(element => {
        // Hydration appends the recovered HTML attribute; its ordering is not
        // document state. Preserve every attribute/value while canonicalizing.
        const canonical = (value: unknown): unknown => {
          if (typeof value === 'string' && value.startsWith('<')) {
            const dom = new DOMParser().parseFromString(value, 'text/html');
            for (const node of dom.querySelectorAll('*')) {
              const attrs = [...node.attributes].map(a => [a.name, a.value]).sort(([a], [b]) => a.localeCompare(b));
              for (const attr of [...node.attributes]) node.removeAttribute(attr.name);
              for (const [name, text] of attrs) node.setAttribute(name, text);
            }
            return dom.body.innerHTML;
          }
          if (Array.isArray(value)) return value.map(canonical);
          if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, canonical(item)]));
          return value;
        };
        return canonical((element as HTMLElement & { editor: { getJSON(): unknown } }).editor.getJSON());
      });
      const characters = () => editor.evaluate(element => {
        type Node = { type: string; text?: string; attrs: Record<string, unknown>; content?: Node[];
          marks?: { type: string; attrs: Record<string, unknown> }[] };
        const doc = (element as HTMLElement & { editor: { getJSON(): Node } }).editor.getJSON();
        return doc.content!.map(p => ({
          paragraph: Object.fromEntries(Object.entries(p.attrs).filter(([key]) => key !== 'sourceParagraph')),
          characters: (p.content || []).flatMap(n => [...(n.type === 'wordTab' ? '\t' : n.text || '')].map(text => ({
            text, marks: (n.marks || []).filter(m => m.type !== 'wordEditRun'),
          }))),
        }));
      });
      await openStory(); const initial = await snapshot();
      if (mode !== 'unchanged') {
        await editor.focus(); await page.keyboard.press('Control+Home');
        if (mode === 'split') { await page.keyboard.press('ArrowRight'); await page.keyboard.press('Enter'); }
        else await page.keyboard.insertText('Q');
        const changed = await snapshot(); expect(changed).not.toEqual(initial);
        await page.keyboard.press('Control+z'); await expect.poll(snapshot).toEqual(initial);
        await page.keyboard.press('Control+y'); await expect.poll(snapshot).toEqual(changed);
      }
      const beforeMigration = await snapshot(), beforeCharacters = await characters();
      if (!body) await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      const stored = (strip = false, fingerprint?: string) => page.evaluate(({ name, strip, fingerprint }) =>
        new Promise<OfficeFile>((resolve, reject) => {
          const request = indexedDB.open('noffice-workspace'); request.onerror = () => reject(request.error);
          request.onsuccess = () => {
            const db = request.result, tx = db.transaction('files', strip || fingerprint ? 'readwrite' : 'readonly');
            const store = tx.objectStore('files'), all = store.getAll(); let file: OfficeFile;
            all.onsuccess = () => {
              file = all.result.find((item: OfficeFile) => item.name === name);
              if (!file || file.content.kind !== 'word') { tx.abort(); return; }
              if (strip) {
                const clean = (html: string) => {
                  const dom = new DOMParser().parseFromString(html, 'text/html');
                  for (const p of dom.querySelectorAll('[data-word-paragraph-script]')) p.removeAttribute('data-word-paragraph-script');
                  return dom.body.innerHTML;
                };
                delete file.content.paragraphScriptsVersion; file.content.html = clean(file.content.html);
                for (const part of file.content.stories?.parts || []) part.html = clean(part.html);
                if (file.content.stories?.emptyTemplates) for (const kind of ['header', 'footer'] as const) {
                  const html = file.content.stories.emptyTemplates[kind];
                  if (html) file.content.stories.emptyTemplates[kind] = clean(html);
                }
              }
              if (fingerprint) file.original!.contentFingerprint = fingerprint;
              if (strip || fingerprint) store.put(file);
            };
            tx.oncomplete = () => { db.close(); resolve(file); };
            tx.onerror = tx.onabort = () => { db.close(); reject(tx.error || Error('Missing saved document')); };
          };
        }), { name, strip, fingerprint });
      const stripped = await stored(true);
      if (mode === 'unchanged') await stored(false, await contentFingerprint(stripped.content));
      const legacy = await stored();
      await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
      await page.getByRole('button', { name, exact: true }).click();
      const outputs: Record<string, string> = {};
      const download = async (label: string, filename: string) => {
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        const pending = page.waitForEvent('download');
        pending.catch(() => {});
        await page.getByRole('button', { name: label, exact: true }).click();
        const downloaded = await Promise.race([pending,
          page.locator('.error-banner').waitFor({ state: 'visible' }).then(async () => {
            throw Error(await page.locator('.error-banner').innerText());
          })]);
        await downloaded.saveAs(root + '/' + filename);
        const bytes = await fs.readFile(root + '/' + filename); outputs[filename] = hash(bytes); return bytes;
      };
      if (mode === 'split') {
        await expect(page.getByRole('region', { name: 'Document recovery' }).getByRole('alert'))
          .toHaveText(/changed paragraph structure with missing script formatting/);
        await expect(page.getByRole('textbox', { name: 'Document text', exact: true })).toHaveCount(0);
        const downloads: string[] = []; page.on('download', value => downloads.push(value.suggestedFilename()));
        for (const label of ['DOCX file Editable in Microsoft Word', 'PDF file', 'Print / Save as PDF']) {
          await page.evaluate(() => { window.print = () => { throw Error('Recovery must not open print'); }; });
          await page.getByRole('button', { name: 'Export', exact: true }).click();
          await page.getByRole('button', { name: label, exact: true }).click();
          await expect(page.locator('.error-banner')).toContainText(/changed paragraph structure with missing script formatting/);
          expect(downloads).toEqual([]); expect(await stored()).toEqual(legacy);
          await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
        }
        const pending = page.waitForEvent('download');
        await page.getByRole('button', { name: 'Download original DOCX', exact: true }).click();
        await (await pending).saveAs(root + '/original.docx');
        expect(await fs.readFile(root + '/original.docx')).toEqual(source);
      } else {
        await openStory(); await expect.poll(snapshot).toEqual(beforeMigration);
        await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.insertText('Z');
        const typed = await snapshot(); expect(typed).not.toEqual(beforeMigration);
        await page.keyboard.press('Control+z'); await expect.poll(snapshot).toEqual(beforeMigration);
        await page.keyboard.press('Control+y'); await expect.poll(snapshot).toEqual(typed);
        await page.keyboard.press('Control+z'); await expect.poll(snapshot).toEqual(beforeMigration);
        if (!body) await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
      }
      const backup = JSON.parse((await download('Noffice backup Full editable model and retained original · .noffice', 'recovered.noffice')).toString());
      expect(Buffer.from(backup.original.base64, 'base64')).toEqual(source);
      if (mode === 'split') { expect(backup.content).toEqual(legacy.content); expect(await stored()).toEqual(legacy); }
      else {
        expect(backup.content.paragraphScriptsVersion).toBe(1);
        const docx = await download('DOCX file Editable in Microsoft Word', 'recovered.docx');
        if (mode === 'unchanged') expect(docx).toEqual(source);
        await download('PDF file', 'recovered.pdf');
        await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/recovered.docx');
        await openStory(); await expect.poll(characters).toEqual(beforeCharacters);
        if (!body) await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
      }
      if (mode === 'split') {
        await page.locator('input[type=file][multiple]').setInputFiles(root + '/original.docx');
        await expect(page.getByRole('region', { name: 'Document recovery' })).toHaveCount(0);
        await expect(page.getByRole('textbox', { name: 'Document text', exact: true })).toBeVisible();
        expect(await download('DOCX file Editable in Microsoft Word', 'reopened-original.docx')).toEqual(source);
        expect(await stored()).toEqual(legacy);
      }
      expect(errors).toEqual([]);
      await fs.writeFile(root + '/browser-report.json', JSON.stringify({ mode, sourceHash: hash(source), outputs, errors,
        buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-script-recovery.spec.ts')),
        referenceHash: hash(await fs.readFile('tests/fixtures/native-word-script-scopes.json')) }, null, 2));
    });
