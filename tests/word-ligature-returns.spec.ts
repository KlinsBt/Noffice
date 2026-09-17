import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-ligatures.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`native ligature ${row.name}: further edit returns through history and reload`, async ({ page }) => {
  test.skip(process.env.NOFFICE_LIGATURE_RETURN !== '1', 'Requires independently edited installed Word output');
  const run = process.env.NOFFICE_LIGATURE_RETURN_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid native return evidence folder');
  const root = `.local/word-ligatures/browser-v2/${row.name}/native-v1`, output = root + '/' + run;
  await fs.mkdir(output, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile(root + '/returned.docx');
  const nativeBytes = await fs.readFile(root + '/native-report.json');
  const native = JSON.parse(nativeBytes.toString().replace(/^\uFEFF/, ''));
  const expected = native.stages.find((stage: { stage: string }) => stage.stage === 'returned');
  expect(hash(source)).toBe(expected.docxHash);
  expect(hash(await fs.readFile(root + '/returned.pdf'))).toBe(expected.pdfHash);
  expect(native.executableHash).toBe(row.executableHash);
  expect(expected.flags).toBe(row.flags); expect(expected.markFlags).toBe(row.markFlags);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: `ligature-return-${row.name}.docx`,
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const footer = page.locator('.word-page-story[data-word-story*="footer"] p').first();
  const check = async () => {
    await expect(page.locator('.section-page')).toHaveCount(2);
    await expect(footer).toHaveText(expected.text);
    await expect(footer).toHaveAttribute('data-word-paragraph-font-features', String(row.markFlags));
    const values = await footer.locator('[data-word-font-features]').evaluateAll(elements =>
      elements.map(element => element.getAttribute('data-word-font-features')));
    expect(values.length).toBeGreaterThan(0); expect(new Set(values)).toEqual(new Set([String(row.flags)]));
  };
  await check(); await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: /^Section 1 \u2014 Default footer/ }).click();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.insertText('X');
  await expect(editor.locator('p').first()).toHaveText('X' + expected.text);
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText(expected.text);
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText('X' + expected.text);
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText(expected.text);
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click(); await check();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `ligature-return-${row.name}`, exact: true }).click(); await check();
  const outputs: Record<string, string> = {};
  for (const [filename, label] of [['returned.docx', 'DOCX file Editable in Microsoft Word'], ['returned.pdf', 'PDF file']]) {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await pending).saveAs(output + '/' + filename);
    outputs[filename] = hash(await fs.readFile(output + '/' + filename));
  }
  expect(await fs.readFile(output + '/returned.docx')).toEqual(source); expect(errors).toEqual([]);
  await fs.writeFile(output + '/browser-report.json', JSON.stringify({ sourceHash: hash(source),
    nativeReceiptHash: hash(nativeBytes), outputs, errors, buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-ligature-returns.spec.ts')) }, null, 2));
});
