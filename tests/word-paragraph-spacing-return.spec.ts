import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-paragraph-spacing.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

test.skip(process.env.NOFFICE_PARAGRAPH_SPACING_RETURN !== '1', 'Requires current independently edited native returns');
for (const row of reference.rows) test(`Word native spacing return ${row.name}: lines and Auto survive native edit and reload`, async ({ page }) => {
  const root = `.local/word-paragraph-spacing-modes/returns-v2/${row.name}`;
  const run = process.env.NOFFICE_PARAGRAPH_RETURN_RUN || 'returns-browser-v2';
  if (!/^returns-browser-v\d+$/.test(run)) throw Error('Invalid native-return evidence folder');
  const output = `.local/word-paragraph-spacing-modes/${run}/${row.name}`;
  await fs.mkdir(output, { recursive: true });
  const source = await fs.readFile(root + '/returned.docx');
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const report = JSON.parse((await fs.readFile('.local/word-paragraph-spacing-modes/returns-v2/native-report.json', 'utf8')).replace(/^\uFEFF/, ''));
  const native = report.rows.find((r: { name: string }) => r.name === row.name);
  expect(hash(source)).toBe(native.docxHash);
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: `returned-${row.name}.docx`, buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  const check = async () => {
    await expect(body.locator(':scope > p').nth(2)).toHaveText('CNative');
    await body.locator(':scope > p').nth(2).click(); await page.getByRole('button', { name: 'Layout', exact: true }).click();
    await expect(page.getByLabel('Before paragraph unit', { exact: true })).toHaveValue('lines');
    await expect(page.getByRole('spinbutton', { name: 'Before paragraph', exact: true })).toHaveValue('0.5');
    await expect(page.getByLabel('After paragraph unit', { exact: true })).toHaveValue('auto');
    await expect(page.getByRole('spinbutton', { name: 'After paragraph', exact: true })).toBeDisabled();
  };
  const outputs: Record<string, string> = {};
  const download = async (name: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download'); await page.getByRole('button', { name: 'PDF file', exact: true }).click();
    await (await pending).saveAs(output + '/' + name); outputs[name] = hash(await fs.readFile(output + '/' + name));
  };
  await check(); await download('returned.pdf');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `returned-${row.name}`, exact: true }).click(); await check(); await download('reloaded.pdf');
  await fs.writeFile(output + '/browser-report.json', JSON.stringify({ sourceHash: hash(source), outputs,
    buildHash: await wordStoryBuildHash(), nativeReceiptHash: hash(await fs.readFile('.local/word-paragraph-spacing-modes/returns-v2/native-report.json')),
    testHash: hash(await fs.readFile('tests/word-paragraph-spacing-return.spec.ts')) }, null, 2));
});
