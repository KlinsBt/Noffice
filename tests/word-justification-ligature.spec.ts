import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-justification-calibri-browser.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const kind of ['body', 'header', 'footer']) test(`Word Calibri ${kind}: individual letters inside ligatures remain editable`, async ({ page }) => {
  const row = reference.rows.find(row => row.name === `regular-${kind}-below`)!;
  const run = process.env.NOFFICE_JUSTIFICATION_LIGATURE || 'ligature-browser-v1';
  expect(run).toMatch(/^ligature-browser-v\d+$/);
  const root = `.local/word-story-paragraph-layout/justification-calibri-eleven/${run}/${kind}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  expect(hash(await fs.readFile(row.fixture))).toBe(row.sourceHash);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(row.fixture);
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
  const word = kind === 'body' ? 'fifty' : kind === 'header' ? 'office' : 'difficult';
  const wordStart = row.text.indexOf(word); expect(wordStart).toBeGreaterThanOrEqual(0);
  const offset = wordStart + (kind === 'body' ? 1 : kind === 'header' ? 2 : 3);
  const locate = async () => {
    await editor.focus(); await page.keyboard.press('Control+Home');
    if (kind !== 'body') {
      await page.keyboard.press('Control+ArrowDown'); await page.keyboard.press('Control+ArrowDown');
      await page.keyboard.press('Home');
    }
    for (let i = 0; i < offset; i++) await page.keyboard.press('ArrowRight');
    expect(await paragraph.evaluate(element => {
      const selection = window.getSelection()!;
      if (!selection.isCollapsed || !selection.anchorNode || !element.contains(selection.anchorNode)) return null;
      const range = document.createRange(); range.selectNodeContents(element);
      range.setEnd(selection.anchorNode, selection.anchorOffset); return range.toString();
    })).toBe(row.text.slice(0, offset));
  };
  const restored = async () => {
    await expect(paragraph).toHaveText(row.text);
    await expect(paragraph).toHaveAttribute('data-word-justification', 'modern');
    await expect(undo).toBeDisabled();
  };
  for (const key of ['Delete', 'Backspace']) {
    await locate(); await page.keyboard.press(key);
    const start = key === 'Delete' ? offset : offset - 1;
    const expected = row.text.slice(0, start) + row.text.slice(start + 1);
    await expect(paragraph).toHaveText(expected);
    await undo.click(); await restored(); await redo.click(); await expect(paragraph).toHaveText(expected);
    await undo.click(); await restored();
  }
  await locate(); await page.keyboard.type('X');
  const finalText = row.text.slice(0, offset) + 'X' + row.text.slice(offset);
  await expect(paragraph).toHaveText(finalText); await undo.click(); await restored();
  await redo.click(); await expect(paragraph).toHaveText(finalText);
  if (kind !== 'body') await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click();
  const saved = kind === 'body' ? paragraph : page.locator('.word-story-measure-host p').filter({ hasText: finalText }).first();
  await expect(saved).toHaveText(finalText); await expect(saved).toHaveAttribute('data-word-justification', 'modern');
  const outputs: Record<string, string> = {};
  for (const [label, file] of [['DOCX file Editable in Microsoft Word', 'edited.docx'], ['PDF file', 'edited.pdf']]) {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click(); await (await pending).saveAs(root + '/' + file);
    outputs[file] = hash(await fs.readFile(root + '/' + file));
  }
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/edited.docx');
  await expect(saved).toHaveText(finalText); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ name: row.name, kind, sourceHash: row.sourceHash,
    errors, outputs, actions: { offset, replacement: 'X' }, finalText,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-justification-ligature.spec.ts')) }, null, 2));
});
