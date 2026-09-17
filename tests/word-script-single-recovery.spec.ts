import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-script-single-lines.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows.filter(row => row.scope === 'all'))
  for (const variant of ['wrapped', 'resource', 'hard-break', 'size20'] as const)
    test(`Word Single script recovery ${row.name} ${variant}`, async ({ page }) => {
      const run = process.env.NOFFICE_SCRIPT_SINGLE_RECOVERY_RUN || 'browser-v1';
      if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
      const root = `.local/word-script-single-lines/recovery/${run}/${row.name}-${variant}`;
      await fs.mkdir(root, { recursive: true });
      const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
      const source = await fs.readFile('tests/fixtures/word-script-single-lines/' + row.sourceFile);
      expect(hash(source)).toBe(row.stages[0].docxHash);
      const name = 'Single recovery ' + row.name + ' ' + variant, body = row.kind === 'Body';
      const kind = row.kind === 'Headers' ? 'header' : 'footer';
      const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
      await page.context().grantPermissions(['local-fonts']); await page.goto('/');
      await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      const open = async () => {
        if (body) return;
        await page.getByRole('button', { name: 'Insert', exact: true }).click();
        await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
        await page.getByRole('button', { name: new RegExp(`^Section 1 — Default ${kind}`) }).click();
      };
      const editor = page.getByRole('textbox', { name: body ? 'Document text' : 'Header or footer text', exact: true });
      const snapshot = () => editor.evaluate(element => (element as HTMLElement & { editor: { getJSON(): unknown } }).editor.getJSON());
      await open(); await expect(editor.locator('p').first()).toHaveText('Ab'); const initial = await snapshot();
      await editor.focus(); await page.keyboard.press('Control+Home');
      if (variant === 'hard-break') { await page.keyboard.press('ArrowRight'); await page.keyboard.press('Shift+Enter'); }
      else {
        await page.keyboard.press('Shift+End');
        if (variant === 'size20') {
          const controls = body ? page : page.getByRole('dialog').last();
          const size = controls.getByRole('spinbutton', { name: 'Font size', exact: true });
          await size.fill('20'); await size.press('Enter');
        } else await page.keyboard.insertText('Ab '.repeat(variant === 'resource' ? 1400 : 100));
      }
      const changed = await snapshot(); expect(changed).not.toEqual(initial);
      await editor.focus(); await page.keyboard.press('Control+z'); await expect.poll(snapshot).toEqual(initial);
      await page.keyboard.press('Control+y'); await expect.poll(snapshot).toEqual(changed);
      if (!body) await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      const outputs: Record<string, string> = {};
      const download = async (file: string, label: string) => {
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        const pending = page.waitForEvent('download');
        await page.getByRole('button', { name: label, exact: true }).click();
        await (await pending).saveAs(root + '/' + file); outputs[file] = hash(await fs.readFile(root + '/' + file));
      };
      const rejectPdf = async () => {
        const unexpected: string[] = [], observe = (file: { suggestedFilename(): string }) => unexpected.push(file.suggestedFilename());
        page.on('download', observe);
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        await page.getByRole('button', { name: 'PDF file', exact: true }).click();
        await expect(page.locator('.error-banner')).toContainText('Export failed:');
        expect(unexpected).toEqual([]); page.off('download', observe);
      };
      await rejectPdf(); await download('unsupported.docx', 'DOCX file Editable in Microsoft Word');
      await download('unsupported.noffice', 'Noffice backup Full editable model and retained original · .noffice');
      const backup = JSON.parse((await fs.readFile(root + '/unsupported.noffice')).toString());
      expect(Buffer.from(backup.original.base64, 'base64')).toEqual(source);
      await page.getByRole('button', { name: 'Home', exact: true }).click();
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await open(); await expect.poll(snapshot).toEqual(initial);
      if (!body) await page.getByRole('button', { name: 'Cancel', exact: true }).click();
      await download('recovered.pdf', 'PDF file');
      await page.getByRole('button', { name: 'Home', exact: true }).click();
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
      await page.getByRole('button', { name, exact: true }).click();
      await open(); await expect.poll(snapshot).toEqual(changed);
      if (!body) await page.getByRole('button', { name: 'Cancel', exact: true }).click();
      await rejectPdf();
      await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/unsupported.docx');
      await open();
      // Import assigns new retained-run/source identities; native properties
      // of this actual exported file are checked independently after reopening.
      const reimported = await snapshot();
      expect(errors).toEqual([]);
      await fs.writeFile(root + '/browser-report.json', JSON.stringify({ outputs, errors, changed, reimported,
        sourceHash: hash(source), buildHash: await wordStoryBuildHash(),
        testHash: hash(await fs.readFile('tests/word-script-single-recovery.spec.ts')),
        referenceHash: hash(await fs.readFile('tests/fixtures/native-word-script-single-lines.json')) }, null, 2));
    });
