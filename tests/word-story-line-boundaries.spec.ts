import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-line-spacing-input.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
const cases = [
  { name: 'minimum-multiple', rule: 'auto', amount: 0.058, line: 14 },
  { name: 'rounded-single', rule: 'auto', amount: 1.0020833333333333, line: 240 },
  { name: 'rounded-exact', rule: 'exact', amount: 0.725, line: 15 },
  { name: 'rounded-minimum', rule: 'atLeast', amount: 0.775, line: 15 },
  { name: 'maximum-exact', rule: 'exact', amount: 1584.00001, line: 31680 },
];
async function openStory(page: Page) {
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  // Sub-point lines overlap adjacent pointer hit boxes. Exercise real keyboard
  // recovery; the recorded pointer failure remains an open pagination case.
  await editor.focus(); await page.keyboard.press('Control+Home');
  await page.keyboard.press('Control+ArrowDown'); await page.keyboard.press('Control+ArrowDown');
  return editor;
}
async function download(page: Page, path: string) {
  await page.getByRole('button', { name: 'Export', exact: true }).click(); const waiting = page.waitForEvent('download');
  await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
  await (await waiting).saveAs(path);
}
for (const context of ['body', 'header']) for (const sample of cases)
test(`Word ${context} ${sample.name}: native entry, rejection, history, reload and actual files`, async ({ page }) => {
  expect(reference.rows.find(r => r.rule === sample.rule && r.amount === sample.amount)?.expected).toEqual({ rule: sample.rule, line: sample.line });
  const run = process.env.NOFFICE_STORY_LINE_BOUNDARIES_RUN || 'boundaries-browser-v1'; expect(run).toMatch(/^boundaries-browser-v\d+$/);
  const root = `.local/word-story-line-authoring/${run}/${context}-${sample.name}`; await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-story-spacing-grid-header-before-3-0.docx'); const name = `Line bounds ${context} ${sample.name}`;
  const errors: string[] = []; page.on('pageerror', e => errors.push(e.message)); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const open = async () => {
    if (context === 'header') return openStory(page);
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true }); await editor.locator('p').first().click(); return editor;
  };
  await open(); const scope = context === 'header' ? page.getByRole('dialog') : page;
  const control = scope.getByRole('combobox', { name: 'Line spacing', exact: true }); const initial = await control.inputValue();
  const custom = async () => {
    await control.selectOption('custom'); const dialog = page.getByRole('dialog');
    return { amount: dialog.getByRole('spinbutton', { name: 'Spacing amount', exact: true }), dialog,
      apply: dialog.getByRole('button', { name: context === 'header' ? 'Apply line spacing' : 'Apply', exact: true }),
      cancel: dialog.getByRole('button', { name: context === 'header' ? 'Cancel line spacing' : 'Cancel', exact: true }) };
  };
  let options = await custom(); await options.dialog.getByRole('combobox', { name: 'Spacing rule', exact: true }).selectOption(sample.rule);
  for (const invalid of ['0', sample.rule === 'auto' ? '0.01' : '0.1', '100000']) {
    await options.amount.fill(invalid); await expect(options.apply).toBeDisabled(); await expect(options.amount).toHaveAttribute('aria-invalid', 'true');
  }
  await options.cancel.click(); await expect(control).toHaveValue(initial);
  options = await custom(); await options.dialog.getByRole('combobox', { name: 'Spacing rule', exact: true }).selectOption(sample.rule);
  await options.amount.fill(String(sample.amount)); await expect(options.apply).toBeEnabled(); await options.apply.click();
  const wanted = sample.rule === 'auto' ? String(sample.line / 240) : `${sample.line / 20}pt`; await expect(control).toHaveValue(wanted);
  options = await custom(); await expect(options.apply).toBeEnabled(); await options.apply.click();
  await scope.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(control).toHaveValue(initial);
  await expect(scope.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
  await scope.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(control).toHaveValue(wanted);
  if (context === 'header') await scope.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await download(page, root + '/edited.docx'); await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click(); await page.getByRole('button', { name, exact: true }).click();
  await open(); await expect(control).toHaveValue(wanted);
  if (context === 'header') await scope.getByRole('button', { name: 'Cancel', exact: true }).click();
  await download(page, root + '/reloaded.docx');
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/edited.docx'); await open(); await expect(control).toHaveValue(wanted);
  if (context === 'header') await scope.getByRole('button', { name: 'Cancel', exact: true }).click(); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ name: `${context}-${sample.name}`, context, sample, sourceHash: hash(source), errors,
    outputs: { 'edited.docx': hash(await fs.readFile(root + '/edited.docx')), 'reloaded.docx': hash(await fs.readFile(root + '/reloaded.docx')) },
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-story-line-boundaries.spec.ts')), referenceHash: hash(await fs.readFile('tests/fixtures/native-word-line-spacing-input.json')) }, null, 2));
});

for (const native of reference.imports) test(`Word imported tiny ${native.sample.rule}-${native.sample.line}: preserve, reject reapply and recover`, async ({ page }) => {
  const { rule, line } = native.sample; const name = `Tiny ${rule} ${line}`;
  const root = `.local/word-story-line-authoring/tiny-browser-v1/${rule}-${line}`; await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile(`tests/fixtures/word-line-spacing-tiny-${rule}-${line}.docx`); expect(hash(source)).toBe(native.sample.sha256);
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await openStory(page); const dialog = page.getByRole('dialog'); const control = dialog.getByRole('combobox', { name: 'Line spacing', exact: true });
  const wanted = rule === 'auto' ? String(line / 240) : `${line / 20}pt`; await expect(control).toHaveValue(wanted); await control.selectOption('custom');
  await expect(dialog.getByRole('button', { name: 'Apply line spacing', exact: true })).toBeDisabled();
  await dialog.getByRole('button', { name: 'Cancel line spacing', exact: true }).click();
  await control.selectOption('1'); await dialog.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(control).toHaveValue(wanted);
  await expect(dialog.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
  await dialog.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(control).toHaveValue('1');
  await dialog.getByRole('button', { name: 'Undo', exact: true }).click(); await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await download(page, root + '/preserved.docx'); expect(hash(await fs.readFile(root + '/preserved.docx'))).toBe(hash(source));
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click(); await page.getByRole('button', { name, exact: true }).click();
  await openStory(page); await expect(control).toHaveValue(wanted); await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(source), preservedHash: hash(await fs.readFile(root + '/preserved.docx')),
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-story-line-boundaries.spec.ts')) }, null, 2));
});
