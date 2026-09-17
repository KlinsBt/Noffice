import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-story-paragraph-layout.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
const controls: Record<string, string> = { LeftIndent: 'Indent before text', RightIndent: 'Indent after text',
  FirstLineIndent: 'First line (negative for hanging)', KeepWithNext: 'Keep with next', KeepTogether: 'Keep lines together',
  PageBreakBefore: 'Page break before', WidowControl: 'Widow/orphan control' };
for (const row of reference.rows) test(`Word ${row.name}: paragraph formatting, history, reload and exported files`, async ({ page }) => {
  const run = process.env.NOFFICE_STORY_PARAGRAPH_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid paragraph-layout output folder');
  const root = `.local/word-story-paragraph-layout/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (value: Buffer) => createHash('sha256').update(value).digest('hex');
  const source = row.states.find(s => s.stage === 'source')!;
  const input = await fs.readFile(source.fixture); expect(hash(input)).toBe(source.docxHash);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  const outputs: Record<string, string> = {};
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  const name = 'Story paragraph ' + row.name;
  await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: input,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect(page.locator('.section-page')).toHaveCount(2);
  const download = async (file: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await waiting).saveAs(root + '/' + file); outputs[file] = hash(await fs.readFile(root + '/' + file));
  };
  const open = async () => {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${row.kind}`) }).click();
    const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    await expect(editor.locator('p')).toHaveCount(5); await editor.locator('p').nth(2).click(); return editor;
  };
  const property = row.requested.property, dialog = page.getByRole('dialog');
  const isIndent = property.endsWith('Indent'), isAlign = property === 'Alignment';
  const originalValue = property === 'WidowControl' ? -1 : 0;
  const alignment = (value: number) => ['left', 'center', 'right', 'justify'][value];
  const verify = async (value: number) => {
    if (isAlign) await expect(dialog.getByRole('button', { name: 'Align ' + alignment(value), exact: true })).toHaveAttribute('aria-pressed', 'true');
    else if (isIndent) await expect(dialog.getByRole('spinbutton', { name: controls[property], exact: true })).toHaveValue(String(value));
    else await expect(dialog.getByRole('checkbox', { name: controls[property], exact: true })).toBeChecked({ checked: value === -1 });
  };
  const apply = async (value: number) => {
    if (isAlign) await dialog.getByRole('button', { name: 'Align ' + alignment(value), exact: true }).click();
    else if (isIndent) {
      const input = dialog.getByRole('spinbutton', { name: controls[property], exact: true });
      await input.fill(String(value)); await input.press('Enter');
    } else await dialog.getByRole('checkbox', { name: controls[property], exact: true }).setChecked(value === -1);
  };
  await download('source.pdf'); let editor = await open(); await verify(originalValue);
  if (isAlign) {
    await editor.focus();
    await page.keyboard.press('Control+Shift+' + ['', 'e', 'r', 'j'][row.requested.value]);
    await verify(row.requested.value);
    await page.keyboard.press('Control+z'); await verify(originalValue);
    await expect(dialog.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
  }
  if (isIndent) {
    const input = dialog.getByRole('spinbutton', { name: controls[property], exact: true });
    for (const invalid of ['', '100000', '-100000']) {
      await input.fill(invalid); await input.press('Enter'); await expect(input).toHaveAttribute('aria-invalid', 'true');
      await expect(dialog.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
      await input.press('Escape'); await verify(originalValue);
    }
  }
  if (row.name === 'header-indent-start') {
    const before = dialog.getByRole('spinbutton', { name: 'Indent before text', exact: true });
    const after = dialog.getByRole('spinbutton', { name: 'Indent after text', exact: true });
    await before.fill('36'); await before.press('Tab'); await expect(after).toBeFocused();
    await verify(36); await after.press('Shift+Tab'); await expect(before).toBeFocused();
    await dialog.getByRole('button', { name: 'Undo', exact: true }).click(); await verify(originalValue);
  }
  await apply(row.requested.value); await verify(row.requested.value);
  // Reapplying a value cannot create another undo event.
  await apply(row.requested.value);
  await dialog.getByRole('button', { name: 'Undo', exact: true }).click(); await verify(originalValue);
  await expect(dialog.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
  await dialog.getByRole('button', { name: 'Redo', exact: true }).click(); await verify(row.requested.value);
  await expect(editor.locator('p')).toHaveCount(5);
  await dialog.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); editor = await open(); await verify(originalValue);
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await page.getByRole('button', { name: 'Home', exact: true }).click(); await page.getByRole('button', { name: 'Redo', exact: true }).click();
  editor = await open(); await verify(row.requested.value); await apply(originalValue);
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  editor = await open(); await verify(row.requested.value); await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(page.locator('.section-page')).toHaveCount(2);
  await download('edited.docx', 'DOCX file Editable in Microsoft Word'); await download('edited.pdf');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click(); await page.getByRole('button', { name, exact: true }).click();
  editor = await open(); await verify(row.requested.value); await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await download('reloaded.docx', 'DOCX file Editable in Microsoft Word'); await download('reloaded.pdf');
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/edited.docx');
  editor = await open(); await verify(row.requested.value); await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await download('reimported.pdf'); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ name: row.name, sourceHash: source.docxHash, outputs, errors,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-story-paragraph-layout.spec.ts')),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-story-paragraph-layout.json')) }, null, 2));
});
