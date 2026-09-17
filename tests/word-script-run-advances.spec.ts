import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import native from './fixtures/native-word-script-run-advances.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

test('Word script strings retain native advances across source run boundaries, editing, print and reload', async ({ page }) => {
  const run = process.env.NOFFICE_SCRIPT_ADVANCE_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
  const root = `.local/word-script-run-advances/${run}`;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/' + native.sourceFile);
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  expect(hash(source)).toBe(native.sourceHash);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles('tests/fixtures/' + native.sourceFile);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const paragraphs = editor.locator('p');
  await expect(paragraphs).toHaveText(native.rows.map(row => row.text));
  await expect(page.locator('.section-page')).toHaveCount(3);
  const model = () => editor.evaluate(el => (el as HTMLElement & { editor: { getJSON(): unknown } }).editor.getJSON());
  const before = await model();
  const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.insertText('Q');
  await expect(paragraphs.first()).toHaveText('QAXb');
  await page.keyboard.press('Control+z'); expect(await model()).toEqual(before);
  await page.keyboard.press('Control+y'); await expect(paragraphs.first()).toHaveText('QAXb');
  await page.keyboard.press('Control+z'); expect(await model()).toEqual(before);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const outputs: Record<string, string> = {};
  const download = async (file: string, label: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await pending).saveAs(root + '/' + file);
    outputs[file] = hash(await fs.readFile(root + '/' + file));
  };
  await download('recovered.docx', 'DOCX file Editable in Microsoft Word');
  expect(outputs['recovered.docx']).toBe(native.sourceHash);
  await download('recovered.pdf', 'PDF file');
  await page.emulateMedia({ media: 'print' });
  await page.pdf({ path: root + '/browser-print.pdf', preferCSSPageSize: true, printBackground: true });
  await page.emulateMedia({ media: 'screen' }); expect(await model()).toEqual(before);
  outputs['browser-print.pdf'] = hash(await fs.readFile(root + '/browser-print.pdf'));
  // Printing must return to the same editable document and history stack.
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.insertText('Q');
  await expect(paragraphs.first()).toHaveText('QAXb');
  await page.keyboard.press('Control+z'); expect(await model()).toEqual(before);
  await page.keyboard.press('Control+y'); await expect(paragraphs.first()).toHaveText('QAXb');
  await page.keyboard.press('Control+z'); expect(await model()).toEqual(before);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click();
  await expect(paragraphs).toHaveText(native.rows.map(row => row.text)); expect(await model()).toEqual(before);
  await download('reloaded.pdf', 'PDF file');
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/recovered.docx');
  await expect(paragraphs).toHaveText(native.rows.map(row => row.text)); expect(await model()).toEqual(before);
  await download('reimported.pdf', 'PDF file'); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ outputs, errors,
    sourceHash: hash(source), buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-script-run-advances.spec.ts')),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-script-run-advances.json')) }, null, 2));
});
