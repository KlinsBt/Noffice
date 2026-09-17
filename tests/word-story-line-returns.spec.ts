import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-story-line-authoring.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
test.skip(process.env.NOFFICE_STORY_LINE_RETURN !== '1', 'Requires independently captured native re-edit files');
for (const row of reference.rows) test(`Word native ${row.name} return retains exact spacing, typing history and original bytes`, async ({ page }) => {
  const sourceRun = process.env.NOFFICE_STORY_LINE_SOURCE_RUN || 'browser-v2';
  const run = process.env.NOFFICE_STORY_LINE_RETURN_RUN || 'returns-browser-v1';
  if (!/^browser-v\d+$/.test(sourceRun) || !/^returns-browser-v\d+$/.test(run)) throw Error('Invalid native return folder');
  const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
  const nativeRoot = '.local/word-story-line-authoring/' + sourceRun;
  const nativeBytes = await fs.readFile(nativeRoot + '/native-report.json');
  const native = JSON.parse(nativeBytes.toString('utf8').replace(/^\uFEFF/, ''));
  const receipt = native.rows.find((r: { name: string }) => r.name === row.name);
  const source = await fs.readFile(nativeRoot + '/' + row.name + '/native-v1/returned.docx'); expect(hash(source)).toBe(receipt.returnedHash);
  const root = '.local/word-story-line-authoring/' + run + '/' + row.name; await fs.mkdir(root, { recursive: true });
  const errors: string[] = []; page.on('pageerror', e => errors.push(e.message)); const outputs: Record<string, string> = {};
  await page.context().grantPermissions(['local-fonts']); await page.goto('/'); const name = 'Native story lines ' + row.name;
  await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect(page.locator('.section-page')).toHaveCount(2);
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${row.kind}`) }).click();
  const dialog = page.getByRole('dialog'), editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await expect(editor.locator('p').nth(2)).toHaveText('Native\tB'); await editor.locator('p').nth(2).click();
  await expect(dialog.getByRole('combobox', { name: 'Line spacing', exact: true })).toHaveValue('40pt');
  await page.keyboard.press('End'); await page.keyboard.insertText('!'); await expect(editor.locator('p').nth(2)).toHaveText('Native\tB!');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').nth(2)).toHaveText('Native\tB');
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').nth(2)).toHaveText('Native\tB!');
  await page.keyboard.press('Control+z'); await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  const download = async (file: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click(); await (await waiting).saveAs(root + '/' + file);
    outputs[file] = hash(await fs.readFile(root + '/' + file));
  };
  await download('original.docx', 'DOCX file Editable in Microsoft Word'); expect(outputs['original.docx']).toBe(hash(source));
  await download('returned.pdf');
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click(); await page.getByRole('button', { name, exact: true }).click();
  await expect(page.locator(`.word-page-story[data-word-story*="${row.kind}"]`).first().locator('p').nth(2)).toHaveText('Native\tB');
  await download('reloaded.pdf'); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(source), sourceRun, outputs, errors,
    nativeReceiptHash: hash(nativeBytes), buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-story-line-returns.spec.ts')) }, null, 2));
});
