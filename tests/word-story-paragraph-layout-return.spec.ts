import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-story-paragraph-layout.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
for (const row of reference.rows) test(`Word native paragraph return ${row.name}: editing, history and original-file recovery`, async ({ page }) => {
  const run = process.env.NOFFICE_STORY_PARAGRAPH_RETURN_RUN || 'returns-browser-v1';
  if (!/^returns-browser-v\d+$/.test(run)) throw Error('Invalid native-return output folder');
  const sourceRun = process.env.NOFFICE_STORY_PARAGRAPH_SOURCE_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(sourceRun)) throw Error('Invalid native source run');
  const root = `.local/word-story-paragraph-layout/${run}/${row.name}`;
  const source = `.local/word-story-paragraph-layout/${sourceRun}/${row.name}/native-v1/actual-returned.docx`;
  await fs.mkdir(root, { recursive: true });
  const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
  const input = await fs.readFile(source);
  const native = JSON.parse(await fs.readFile(`.local/word-story-paragraph-layout/${sourceRun}/native-report.json`, 'utf8').then(s => s.replace(/^\uFEFF/, '')));
  expect(hash(input)).toBe(native.rows.find((r: { name: string }) => r.name === row.name).states.find((s: { stage: string }) => s.stage === 'actual-returned').actualHash);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  const name = 'Native paragraph ' + row.name;
  await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: input,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect(page.locator('.section-page')).toHaveCount(2);
  const dialog = page.getByRole('dialog');
  const open = async () => {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${row.kind}`) }).click();
    const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    await editor.focus(); await page.keyboard.press('Control+Home');
    await page.keyboard.press('Control+ArrowDown'); await page.keyboard.press('Control+ArrowDown');
    await expect(editor.locator('p').nth(2)).toContainText('Native beta');
    await expect(dialog.getByRole('button', { name: 'Align left', exact: true })).toHaveAttribute('aria-pressed', 'true');
    await expect(dialog.getByRole('spinbutton', { name: 'Indent before text', exact: true })).toHaveValue('0');
    return editor;
  };
  let editor = await open();
  await page.keyboard.press('Control+Shift+ArrowRight'); await page.keyboard.insertText('Again ');
  await expect(editor.locator('p').nth(2)).toContainText('Again beta');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').nth(2)).toContainText('Native beta');
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').nth(2)).toContainText('Again beta');
  await dialog.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await page.getByRole('button', { name: 'Home', exact: true }).click(); await page.getByRole('button', { name: 'Undo', exact: true }).click();
  editor = await open(); await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  const outputs: Record<string, string> = {};
  const download = async (file: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click(); await (await waiting).saveAs(root + '/' + file);
    outputs[file] = hash(await fs.readFile(root + '/' + file));
  };
  await download('recovered.docx', 'DOCX file Editable in Microsoft Word'); expect(outputs['recovered.docx']).toBe(hash(input));
  await download('recovered.pdf'); await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click(); editor = await open();
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click(); await download('reloaded.pdf');
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ name: row.name, source, sourceHash: hash(input), outputs, errors,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-story-paragraph-layout-return.spec.ts')) }, null, 2));
});
