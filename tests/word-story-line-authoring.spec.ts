import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-story-line-authoring.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word ${row.name}: line spacing, inner/outer history, cancellation and actual exports`, async ({ page }) => {
  const run = process.env.NOFFICE_STORY_LINE_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid story line-spacing output folder');
  const root = `.local/word-story-line-authoring/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
  const input = await fs.readFile(row.source); expect(hash(input)).toBe(row.sourceHash);
  const errors: string[] = []; page.on('pageerror', e => errors.push(e.message));
  const outputs: Record<string, string> = {};
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  const name = 'Story lines ' + row.name;
  await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: input,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect(page.locator('.section-page')).toHaveCount(2);
  const download = async (file: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await waiting).saveAs(root + '/' + file); outputs[file] = hash(await fs.readFile(root + '/' + file));
  };
  await download('source.pdf');
  const open = async () => {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${row.kind}`) }).click();
    const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    await expect(editor.locator('p')).toHaveCount(5); await editor.locator('p').nth(2).click(); return editor;
  };
  let editor = await open(); const dialog = page.getByRole('dialog');
  const control = dialog.getByRole('combobox', { name: 'Line spacing', exact: true });
  await expect(control).toHaveValue('40pt');
  await control.selectOption('custom');
  const amount = dialog.getByRole('spinbutton', { name: 'Spacing amount', exact: true });
  await amount.fill('0'); await expect(dialog.getByRole('button', { name: 'Apply line spacing', exact: true })).toBeDisabled();
  await dialog.getByRole('button', { name: 'Cancel line spacing', exact: true }).click(); await expect(control).toHaveValue('40pt');
  const wanted = row.requested.name === 'single' ? '1' : row.requested.name === 'double' ? '2' : row.requested.line + 'pt';
  if (row.requested.name === 'single' || row.requested.name === 'double') await control.selectOption(wanted);
  else {
    await control.selectOption('custom'); await dialog.getByRole('combobox', { name: 'Spacing rule', exact: true }).selectOption(row.requested.name === 'minimum50' ? 'atLeast' : 'exact');
    await amount.fill(String(row.requested.line)); await dialog.getByRole('button', { name: 'Apply line spacing', exact: true }).click();
  }
  await expect(control).toHaveValue(wanted);
  await dialog.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(control).toHaveValue('40pt');
  await expect(dialog.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
  await dialog.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(control).toHaveValue(wanted);
  if (row.name === 'header-single') {
    await editor.focus(); await page.keyboard.press('Control+a'); await expect(control).toHaveValue('');
    await control.selectOption('custom');
    await dialog.getByRole('button', { name: 'Cancel line spacing', exact: true }).click(); await expect(control).toHaveValue('');
    await control.selectOption('2'); await expect(control).toHaveValue('2');
    await dialog.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(control).toHaveValue('');
    await dialog.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(control).toHaveValue('2');
    await dialog.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(control).toHaveValue('');
    await editor.locator('p').nth(2).click(); await expect(control).toHaveValue('1');
    await control.selectOption('1');
    await dialog.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(control).toHaveValue('40pt');
    await expect(dialog.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
    await dialog.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(control).toHaveValue('1');
  }
  await editor.locator('p').nth(2).click(); await page.keyboard.press('Home'); await page.keyboard.press('Shift+ArrowRight'); await page.keyboard.insertText('Edited');
  await expect(editor.locator('p').nth(2)).toHaveText('Edited\tB');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').nth(2)).toHaveText('A\tB');
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').nth(2)).toHaveText('Edited\tB');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').nth(2)).toHaveText('A\tB');
  await dialog.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); editor = await open(); await expect(control).toHaveValue('40pt');
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await page.getByRole('button', { name: 'Home', exact: true }).click(); await page.getByRole('button', { name: 'Redo', exact: true }).click();
  editor = await open(); await expect(control).toHaveValue(wanted);
  await control.selectOption('3'); await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  editor = await open(); await expect(control).toHaveValue(wanted); await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await download('edited.docx', 'DOCX file Editable in Microsoft Word'); await download('edited.pdf');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click(); await page.getByRole('button', { name, exact: true }).click();
  editor = await open(); await expect(control).toHaveValue(wanted); await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await download('reloaded.pdf');
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/edited.docx');
  editor = await open(); await expect(control).toHaveValue(wanted); await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(page.locator('.section-page')).toHaveCount(2); await download('reimported.pdf');
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ name: row.name, sourceHash: row.sourceHash, outputs, errors,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-story-line-authoring.spec.ts')),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-story-line-authoring.json')) }, null, 2));
});
