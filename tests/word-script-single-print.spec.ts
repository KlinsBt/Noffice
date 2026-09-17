import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-script-single-lines.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word Single script print ${row.name}`, async ({ page }) => {
  test.skip(process.env.NOFFICE_SCRIPT_SINGLE_PRINT !== '1', 'Requires verified Single-script authoring downloads');
  const run = process.env.NOFFICE_SCRIPT_SINGLE_SOURCE_RUN || 'browser-v2';
  const printRun = process.env.NOFFICE_SCRIPT_SINGLE_PRINT_RUN || 'print-v2';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
  if (!/^print-v\d+$/.test(printRun)) throw Error('Invalid print evidence folder');
  const inputRoot = `.local/word-script-single-lines/${run}/${row.name}`, root = inputRoot + '/' + printRun;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const input = await fs.readFile(inputRoot + '/recovered.docx');
  const priorBytes = await fs.readFile(inputRoot + '/browser-report.json'), prior = JSON.parse(priorBytes.toString());
  expect(hash(input)).toBe(prior.outputs['recovered.docx']);
  expect(prior.buildHash).toBe(await wordStoryBuildHash());
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(inputRoot + '/recovered.docx');
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(editor).toBeVisible();
  if (row.kind === 'Body')
    await expect(editor.locator('p[data-word-measured-leading]')).toHaveCount(row.stages[0].paragraphs.length);
  else await expect(page.locator('.section-page')).toHaveCount(row.stages[0].pages);
  const model = () => editor.evaluate(el => (el as HTMLElement & { editor: { getJSON(): unknown } }).editor.getJSON());
  const before = await model();
  await page.emulateMedia({ media: 'print' });
  await page.pdf({ path: root + '/browser-print.pdf', preferCSSPageSize: true, printBackground: true });
  await page.emulateMedia({ media: 'screen' });
  expect(await model()).toEqual(before); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ errors,
    outputs: { 'browser-print.pdf': hash(await fs.readFile(root + '/browser-print.pdf')) }, sourceHash: hash(input),
    sourceReceiptHash: hash(priorBytes), buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-script-single-print.spec.ts')),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-script-single-lines.json')) }, null, 2));
});
