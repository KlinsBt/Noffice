import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import arial from './fixtures/native-word-justification-browser.json' with { type: 'json' };
import calibri from './fixtures/native-word-justification-calibri-browser.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
const profile = process.env.NOFFICE_JUSTIFICATION_PROFILE || 'arial-ten';
if (!['arial-ten', 'calibri-eleven'].includes(profile)) throw Error('Invalid justification profile');
const reference = profile === 'calibri-eleven' ? calibri : arial;

for (const kind of ['body', 'header', 'footer']) test(`Word justification ${kind}: caret edits and clipboard paste at visual breaks`, async ({ page }) => {
  const row = reference.rows.find(row => row.name === `regular-${kind}-below`)!;
  const run = process.env.NOFFICE_JUSTIFICATION_CLIPBOARD || 'clipboard-browser-v1';
  expect(run).toMatch(/^clipboard-browser-v\d+$/);
  const root = `.local/word-story-paragraph-layout/justification-${profile}/${run}/${kind}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts', 'clipboard-read', 'clipboard-write']);
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(row.fixture);
  await expect(page.locator('.section-page')).toHaveCount(row.pages);
  let editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  if (kind !== 'body') {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: new RegExp('^Section 1 \\u2014 Default ' + kind) }).click();
    editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  }
  const paragraph = editor.locator('p').nth(kind === 'body' ? 0 : 2);
  const toolbar = kind === 'body' ? page : page.getByRole('dialog');
  const undo = toolbar.getByRole('button', { name: 'Undo', exact: true });
  const redo = toolbar.getByRole('button', { name: 'Redo', exact: true });
  const start = async (offset = 0) => {
    await editor.focus(); await page.keyboard.press('Control+Home');
    if (kind !== 'body') {
      await page.keyboard.press('Control+ArrowDown'); await page.keyboard.press('Control+ArrowDown');
      await page.keyboard.press('Home');
    }
    for (let i = 0; i < offset; i++) await page.keyboard.press('ArrowRight');
  };
  const recovered = async () => {
    await expect(paragraph).toHaveText(row.text);
    await expect(paragraph).toHaveAttribute('data-word-justification', 'modern');
    await expect(undo).toBeDisabled();
  };
  await recovered();
  const edge = row.ends[0];
  expect(row.text[edge]).toBe(' ');
  const joined = row.text.slice(0, edge) + row.text.slice(edge + 1);
  for (const [offset, key] of [[edge + 1, 'Backspace'], [edge, 'Delete']] as const) {
    await start(offset); await page.keyboard.press(key); await expect(paragraph).toHaveText(joined);
    await undo.click(); await recovered(); await redo.click(); await expect(paragraph).toHaveText(joined);
    await undo.click(); await recovered();
  }
  await start(edge + 1); await page.keyboard.type('X');
  await expect(paragraph).toHaveText(row.text.slice(0, edge + 1) + 'X' + row.text.slice(edge + 1));
  await undo.click(); await recovered();
  const count = await editor.locator('p').count();
  await start(edge + 1); await page.keyboard.press('Enter');
  await expect(editor.locator('p')).toHaveCount(count + 1);
  await expect(paragraph).toHaveText(row.text.slice(0, edge + 1));
  await expect(editor.locator('p').nth(kind === 'body' ? 1 : 3)).toHaveText(row.text.slice(edge + 1));
  await undo.click(); await recovered(); await expect(editor.locator('p')).toHaveCount(count);
  await start(edge + 1); await page.keyboard.press('Shift+Enter');
  await expect(paragraph.locator('br:not([data-word-justification-break]):not(.ProseMirror-trailingBreak)')).toHaveCount(1);
  await expect(paragraph).toHaveText(row.text);
  await undo.click(); await recovered();

  if (kind !== 'body') await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const originalDownload = page.waitForEvent('download');
  await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
  await (await originalDownload).saveAs(root + '/recovered.docx');
  expect(hash(await fs.readFile(root + '/recovered.docx'))).toBe(row.sourceHash);
  if (kind !== 'body') {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: new RegExp('^Section 1 \\u2014 Default ' + kind) }).click();
  }

  await start(); const selectedLength = edge + 5;
  for (let i = 0; i < selectedLength; i++) await page.keyboard.press('Shift+ArrowRight');
  await page.evaluate(() => navigator.clipboard.writeText('Replacement'));
  await page.keyboard.press('Control+v');
  const finalText = 'Replacement' + row.text.slice(selectedLength);
  await expect(paragraph).toHaveText(finalText);
  await undo.click(); await recovered(); await redo.click(); await expect(paragraph).toHaveText(finalText);
  if (kind !== 'body') await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click();
  const savedParagraph = kind === 'body' ? paragraph : page.locator('.word-story-measure-host p').filter({ hasText: finalText }).first();
  await expect(savedParagraph).toHaveText(finalText);
  const outputs: Record<string, string> = {};
  for (const [label, file] of [['DOCX file Editable in Microsoft Word', 'edited.docx'], ['PDF file', 'edited.pdf']]) {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click(); await (await waiting).saveAs(root + '/' + file);
    outputs[file] = hash(await fs.readFile(root + '/' + file));
  }
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/edited.docx');
  await expect(savedParagraph).toHaveText(finalText); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ name: row.name, kind, sourceHash: row.sourceHash, errors,
    outputs, actions: { selectedLength, replacement: 'Replacement' }, finalText,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-justification-clipboard.spec.ts')) }, null, 2));
});
