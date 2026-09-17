import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-script-fonts.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word scripts ${row.name}: screen and actual browser print`, async ({ page }) => {
  test.skip(process.env.NOFFICE_SCRIPT_FONT_PRINT !== '1', 'Requires verified script-format authoring downloads');
  const run = process.env.NOFFICE_SCRIPT_FONT_SOURCE_RUN || 'browser-v5';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
  const inputRoot = `.local/word-script-fonts/${run}/${row.name}`;
  const printRun = process.env.NOFFICE_SCRIPT_FONT_PRINT_RUN || 'print-v1';
  if (!/^print-v\d+$/.test(printRun)) throw Error('Invalid print evidence folder');
  const root = inputRoot + '/' + printRun; await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const input = await fs.readFile(inputRoot + '/edited.docx');
  const prior = JSON.parse((await fs.readFile(inputRoot + '/browser-report.json')).toString());
  expect(hash(input)).toBe(prior.outputs['edited.docx']);
  await page.setViewportSize({ width: 1500, height: 1200 });
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(inputRoot + '/edited.docx');
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const paint = `[data-word-script-paint="${row.script.toLowerCase()}"]`;
  // Baseline paint can have a zero-height inline box. Select its visible
  // paragraph instead of a hidden measurement copy or the inline box itself.
  const paragraph = page.locator('p:visible').filter({ has: page.locator(paint) }).first();
  const script = paragraph.locator(paint).first();
  await expect(script).toHaveText('B');
  await expect.poll(() => script.evaluate(el => parseFloat(getComputedStyle(el).fontSize) * .75))
    .toBeCloseTo(row.geometry.b.size, 3);
  await paragraph.scrollIntoViewIfNeeded();
  await paragraph.screenshot({ path: root + '/screen.png' });
  const model = () => editor.evaluate(el => (el as HTMLElement & { editor: { getJSON(): unknown } }).editor.getJSON());
  const before = await model();
  await page.emulateMedia({ media: 'print' });
  await page.pdf({ path: root + '/browser-print.pdf', preferCSSPageSize: true, printBackground: true });
  await page.emulateMedia({ media: 'screen' });
  expect(await model()).toEqual(before); await expect(script).toHaveText('B');
  const outputs: Record<string, string> = {};
  for (const name of ['screen.png', 'browser-print.pdf']) outputs[name] = hash(await fs.readFile(root + '/' + name));
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ outputs, sourceHash: hash(input),
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-script-font-print.spec.ts')),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-script-fonts.json')) }, null, 2));
});
