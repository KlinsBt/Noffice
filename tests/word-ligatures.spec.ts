import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-ligatures.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word ligatures ${row.name}: editing, both history levels and actual exports`, async ({ page }) => {
  const run = process.env.NOFFICE_LIGATURE_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid ligature evidence folder');
  const root = `.local/word-ligatures/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile(`tests/fixtures/word-ligatures-${row.name}.docx`);
  expect(hash(source)).toBe(row.docxHash);
  const sourceText = 'office affinity Native ffi fi fl ft tt ti\tB';
  const editedText = 'affinity affinity Native ffi fi fl ft tt ti\tB';
  const errors: string[] = []; page.on('pageerror', (error) => errors.push(error.message));
  const outputs: Record<string, string> = {};
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: `ligatures-${row.name}.docx`,
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const footer = page.locator('.word-page-story[data-word-story*="footer"] p').first();
  const check = async (text: string) => {
    await expect(page.locator('.section-page')).toHaveCount(2);
    await expect(footer).toHaveText(text);
    await expect(footer).toHaveAttribute('data-word-paragraph-font-features', String(row.markFlags));
    const values = await footer.locator('[data-word-font-features]').evaluateAll((elements) =>
      elements.map((element) => element.getAttribute('data-word-font-features')));
    expect(values.length).toBeGreaterThan(0); expect(new Set(values)).toEqual(new Set([String(row.flags)]));
  };
  const download = async (name: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await pending).saveAs(root + '/' + name);
    outputs[name] = hash(await fs.readFile(root + '/' + name));
  };
  await check(sourceText); await download('source.pdf');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: /^Section 1 \u2014 Default footer/ }).click();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await editor.focus(); await page.keyboard.press('Control+Home');
  for (let i = 0; i < 6; i++) await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.insertText('affinity'); await expect(editor.locator('p').first()).toHaveText(editedText);
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText(sourceText);
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText(editedText);
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click(); await check(editedText);
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); await check(sourceText);
  await page.getByRole('button', { name: 'Redo', exact: true }).click(); await check(editedText);
  await download('edited.pdf'); await download('edited.docx', 'DOCX file Editable in Microsoft Word');
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `ligatures-${row.name}`, exact: true }).click();
  await check(editedText); await download('reloaded.pdf');
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: `ligature-export-${row.name}.docx`,
    buffer: await fs.readFile(root + '/edited.docx'), mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await check(editedText);
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: row.docxHash, outputs,
    flags: row.flags, markFlags: row.markFlags, errors, buildHash: await wordStoryBuildHash(),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-ligatures.json')),
    testHash: hash(await fs.readFile('tests/word-ligatures.spec.ts')) }, null, 2));
});
