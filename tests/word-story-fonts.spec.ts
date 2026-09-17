import { test, expect, type Locator } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-story-fonts.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word story fonts ${row.name}: selection, invalid input, history and exported files`, async ({ page }) => {
  const run = process.env.NOFFICE_STORY_FONT_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid font evidence folder');
  const root = `.local/word-story-fonts/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-tab-leader-story-arial10.docx');
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  expect(hash(source)).toBe(row.sourceHash);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: `Story fonts ${row.name}.docx`,
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const story = page.locator(`.word-page-story[data-word-story*="${row.kind}"]`).first();
  const styles = async (element: Locator, family: string, points: number) => {
    let values: { family: string; points: number }[] = [];
    await expect.poll(async () => {
      values = await element.evaluate(root => {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT), result: { family: string; points: number }[] = [];
      while (walker.nextNode()) {
        const node = walker.currentNode as Text;
        if (!node.data || node.parentElement!.closest('.ProseMirror-widget,[contenteditable=false]')) continue;
        const style = getComputedStyle(node.parentElement!);
        result.push({ family: style.fontFamily.replace(/^['"]|['"]$/g, ''), points: parseFloat(style.fontSize) * .75 });
      }
      return result;
      });
      return values.length > 0 && values.every(value =>
        value.family === family && Math.abs(value.points - points) < .001);
    }, { message: `Story typography should settle to ${family} ${points}pt` }).toBe(true);
    expect(values.length).toBeGreaterThan(0);
    expect(new Set(values.map(value => value.family))).toEqual(new Set([family]));
    expect(Math.max(...values.map(value => Math.abs(value.points - points)))).toBeLessThan(.001);
  };
  const expected = row.paragraphs[0];
  await expect(page.locator('.section-page')).toHaveCount(2);
  await styles(story, 'Arial', 10);
  const open = async () => {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${row.kind}`) }).click();
  };
  const dialog = page.getByRole('dialog').last();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  const format = async () => {
    await editor.focus(); await page.keyboard.press('Control+a');
    if (row.mode === 'arial20') {
      await dialog.getByRole('spinbutton', { name: 'Font size', exact: true }).fill('20');
      await page.keyboard.press('Tab');
    } else await dialog.getByRole('combobox', { name: 'Font family', exact: true }).selectOption('Times New Roman');
    await styles(editor, expected.family, expected.size);
  };
  await open(); await format();
  await dialog.getByRole('button', { name: 'Undo', exact: true }).click(); await styles(editor, 'Arial', 10);
  await dialog.getByRole('button', { name: 'Redo', exact: true }).click(); await styles(editor, expected.family, expected.size);
  for (const invalid of ['0', '20.1', '1638.5']) {
    await dialog.getByRole('spinbutton', { name: 'Font size', exact: true }).fill(invalid); await page.keyboard.press('Tab');
    await expect(dialog.getByRole('spinbutton', { name: 'Font size', exact: true })).toHaveValue(String(expected.size));
    await styles(editor, expected.family, expected.size);
  }
  await dialog.getByRole('button', { name: 'Grow font', exact: true }).click();
  await styles(editor, expected.family, expected.size === 20 ? 22 : 11);
  await dialog.getByRole('button', { name: 'Shrink font', exact: true }).click(); await styles(editor, expected.family, expected.size);
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click(); await styles(story, 'Arial', 10);
  await open(); await format();
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.insertText('Edited'); await expect(editor.locator('p').first()).toHaveText('Edited\tB');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText('A\tB');
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText('Edited\tB');
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await styles(story, expected.family, expected.size); await expect(story.locator('p').first()).toHaveText('Edited\tB');
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); await styles(story, 'Arial', 10);
  await page.getByRole('button', { name: 'Redo', exact: true }).click(); await styles(story, expected.family, expected.size);
  const outputs: Record<string, string> = {};
  const download = async (filename: string, label: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click(); await (await pending).saveAs(root + '/' + filename);
    outputs[filename] = hash(await fs.readFile(root + '/' + filename));
  };
  await download('edited.docx', 'DOCX file Editable in Microsoft Word'); await download('edited.pdf', 'PDF file');
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `Story fonts ${row.name}`, exact: true }).click();
  await styles(story, expected.family, expected.size); await expect(story.locator('p').first()).toHaveText('Edited\tB');
  await download('reloaded.pdf', 'PDF file');
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: row.name + '-export.docx',
    buffer: await fs.readFile(root + '/edited.docx'), mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await styles(story, expected.family, expected.size); await expect(story.locator('p').first()).toHaveText('Edited\tB');
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ outputs, errors, sourceHash: hash(source),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-story-fonts.json')), buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-story-fonts.spec.ts')) }, null, 2));
});
