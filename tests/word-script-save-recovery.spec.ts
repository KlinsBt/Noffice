import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';
import reference from './fixtures/native-word-script-scopes.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows.filter(row => row.scope === 'text-selection'))
  for (const failure of ['late-abort', 'other-tab'] as const)
    test(`Word script save ${row.name} ${failure}: preserve unsaved formatting and newer revisions`, async ({ page, context }) => {
      const run = process.env.NOFFICE_SCRIPT_SAVE_RECOVERY_RUN || 'browser-v1';
      if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
      const root = `.local/word-script-scopes/save-recovery/${run}/${row.name}-${failure}`;
      await fs.mkdir(root, { recursive: true });
      const sourceRow = reference.sources.find(source => source.name === row.source)!;
      const source = await fs.readFile('tests/fixtures/word-script-scopes/' + sourceRow.file);
      const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
      expect(hash(source)).toBe(sourceRow.docxHash);
      const name = `Script save ${row.name}`, body = row.kind === 'Body', kind = row.kind === 'Headers' ? 'header' : 'footer';
      const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
      await context.grantPermissions(['local-fonts']); await page.goto('/');
      await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      const stored = () => page.evaluate(name => new Promise<OfficeFile>((resolve, reject) => {
        const request = indexedDB.open('noffice-workspace'); request.onerror = () => reject(request.error);
        request.onsuccess = () => {
          const db = request.result, tx = db.transaction('files'), all = tx.objectStore('files').getAll(); let file: OfficeFile;
          all.onsuccess = () => {
            file = all.result.find((item: OfficeFile) => item.name === name || item.name === name + ' newer');
            if (!file) tx.abort();
          };
          tx.oncomplete = () => { db.close(); resolve(file); };
          tx.onerror = tx.onabort = () => { db.close(); reject(tx.error); };
        };
      }), name);
      const original = await stored(); let other: Page | undefined;
      if (failure === 'late-abort') await page.evaluate(() => {
        const put = IDBObjectStore.prototype.put;
        IDBObjectStore.prototype.put = function (...args: Parameters<typeof put>) {
          const request = put.apply(this, args), tx = this.transaction;
          if (this.name === 'files') request.addEventListener('success', () => tx.abort(), { once: true });
          return request;
        };
        Object.assign(window, { restoreScriptSave: () => { IDBObjectStore.prototype.put = put; } });
      });
      else {
        other = await context.newPage(); await other.goto('/');
        await other.getByRole('button', { name: 'Recent files', exact: true }).click();
        await other.getByRole('button', { name, exact: true }).click();
        await other.getByRole('textbox', { name: 'File name', exact: true }).fill(name + ' newer');
        await expect(other.getByText('Saved on this device', { exact: true })).toBeVisible();
      }
      const persisted = await stored();
      const openStory = async () => {
        if (body) return;
        await page.getByRole('button', { name: 'Insert', exact: true }).click();
        await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
        await page.getByRole('button', { name: new RegExp(`^Section 1 — Default ${kind}`) }).click();
      };
      const editor = page.getByRole('textbox', { name: body ? 'Document text' : 'Header or footer text', exact: true });
      const characters = () => editor.evaluate(element => {
        type Node = { type: string; text?: string; attrs: Record<string, unknown>; content?: Node[];
          marks?: { type: string; attrs: Record<string, unknown> }[] };
        const p = (element as HTMLElement & { editor: { getJSON(): Node } }).editor.getJSON().content![0];
        const read = (text: string, marks: Node['marks'] = [], paragraph = false) => {
          const attrs = marks?.find(m => m.type === 'textStyle')?.attrs;
          return { text, family: String(attrs?.fontFamily || p.attrs.paragraphFontFamily).replace(/^['"]|['"]$/g, ''),
            size: parseFloat(String(attrs?.fontSize || p.attrs.paragraphFontSize)),
            superscript: (paragraph ? p.attrs.paragraphScript === 'superscript' : marks?.some(m => m.type === 'superscript')) ? -1 : 0,
            subscript: (paragraph ? p.attrs.paragraphScript === 'subscript' : marks?.some(m => m.type === 'subscript')) ? -1 : 0 };
        };
        return [...(p.content || []).flatMap(n => [...(n.type === 'wordTab' ? '\t' : n.text || '')].map(c => read(c, n.marks))), read('\r', [], true)];
      });
      await openStory(); await expect.poll(characters).toEqual(sourceRow.characters);
      await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.down('Shift');
      for (let i = 0; i < sourceRow.characters.length - 1; i++) await page.keyboard.press('ArrowRight');
      await page.keyboard.up('Shift');
      const controls = body ? page : page.getByRole('dialog').last();
      await controls.getByRole('button', { name: row.script, exact: true }).click();
      await expect.poll(characters).toEqual(row.formatted);
      await controls.getByRole('button', { name: 'Undo', exact: true }).click(); await expect.poll(characters).toEqual(sourceRow.characters);
      await controls.getByRole('button', { name: 'Redo', exact: true }).click(); await expect.poll(characters).toEqual(row.formatted);
      if (!body) await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
      await expect(page.getByText('Not saved', { exact: true })).toBeVisible();
      await expect(page.getByRole('alert')).toBeVisible();
      if (failure === 'other-tab') await expect(page.getByRole('alert')).toContainText('another tab');
      expect(await stored()).toEqual(persisted);
      const outputs: Record<string, string> = {};
      const download = async (label: string, file: string) => {
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        const pending = page.waitForEvent('download'); pending.catch(() => {});
        await page.getByRole('button', { name: label, exact: true }).click();
        await (await pending).saveAs(root + '/' + file); const bytes = await fs.readFile(root + '/' + file);
        outputs[file] = hash(bytes); return bytes;
      };
      await download('DOCX file Editable in Microsoft Word', 'unsaved.docx'); await download('PDF file', 'unsaved.pdf');
      const backup = JSON.parse((await download('Noffice backup Full editable model and retained original · .noffice', 'unsaved.noffice')).toString());
      expect(backup.content.paragraphScriptsVersion).toBe(1); expect(Buffer.from(backup.original.base64, 'base64')).toEqual(source);
      expect(await stored()).toEqual(persisted);
      if (failure === 'late-abort') {
        await page.evaluate(() => (window as unknown as { restoreScriptSave(): void }).restoreScriptSave());
        await page.keyboard.press('Control+s'); await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
        await expect(page.getByRole('alert')).toHaveCount(0); expect((await stored()).revision).toBe(original.revision + 1);
      }
      await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
      await page.getByRole('button', { name: name + (failure === 'other-tab' ? ' newer' : ''), exact: true }).click();
      await openStory(); await expect.poll(characters).toEqual(failure === 'late-abort' ? row.formatted : sourceRow.characters);
      if (!body) await page.getByRole('button', { name: 'Cancel', exact: true }).click();
      await page.locator('input[type=file][multiple]').setInputFiles(root + '/unsaved.docx');
      await openStory(); await expect.poll(characters).toEqual(row.formatted);
      if (other) await other.close(); expect(errors).toEqual([]);
      await fs.writeFile(root + '/browser-report.json', JSON.stringify({ failure, sourceHash: hash(source), outputs, errors,
        buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-script-save-recovery.spec.ts')),
        referenceHash: hash(await fs.readFile('tests/fixtures/native-word-script-scopes.json')) }, null, 2));
    });
