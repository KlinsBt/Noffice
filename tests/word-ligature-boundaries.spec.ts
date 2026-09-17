import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-ligature-boundaries.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word ligature boundary ${row.name}: actual PDFs and edit-session recovery`, async ({ page }) => {
  const run = process.env.NOFFICE_LIGATURE_BOUNDARY_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid boundary evidence folder');
  const root = `.local/word-ligatures/boundaries-v1/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile(`tests/fixtures/word-ligature-boundary-${row.name}.docx`);
  expect(hash(source)).toBe(row.docxHash);
  const errors: string[] = []; page.on('pageerror', (error) => errors.push(error.message));
  const outputs: Record<string, string> = {};
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: `boundary-${row.name}.docx`,
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const footer = page.locator('.word-page-story[data-word-story*="footer"] p').first();
  const check = async () => {
    await expect(page.locator('.section-page')).toHaveCount(2);
    await expect(footer).toHaveText('ti\tB');
  };
  const download = async (name: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await pending).saveAs(root + '/' + name);
    outputs[name] = hash(await fs.readFile(root + '/' + name));
  };
  await check(); await download('source.pdf');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: /^Section 1 \u2014 Default footer/ }).click();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await editor.focus(); await page.keyboard.press('Control+Home');
  // Create a real typing-session boundary inside the native cluster. Undo must
  // recover the original styles before the applied document history is tested.
  await page.keyboard.press('ArrowRight'); await page.keyboard.insertText('X');
  await expect(editor.locator('p').first()).toHaveText('tXi\tB');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText('ti\tB');
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText('tXi\tB');
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await expect(footer).toHaveText('tXi\tB');
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); await check();
  await page.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(footer).toHaveText('tXi\tB');
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); await check();
  await download('restored.pdf'); await download('restored.docx', 'DOCX file Editable in Microsoft Word');
  expect(await fs.readFile(root + '/restored.docx')).toEqual(source);
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `boundary-${row.name}`, exact: true }).click();
  await check(); await download('reloaded.pdf');
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: row.docxHash, outputs,
    errors, buildHash: await wordStoryBuildHash(),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-ligature-boundaries.json')),
    testHash: hash(await fs.readFile('tests/word-ligature-boundaries.spec.ts')) }, null, 2));
});
