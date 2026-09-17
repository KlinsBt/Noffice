import { test, expect, type Locator } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-story-font-ranges.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word story font range ${row.name}: native selection properties and history`, async ({ page }) => {
  const run = process.env.NOFFICE_STORY_FONT_RANGE_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid font range evidence folder');
  const root = `.local/word-story-font-ranges/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-tab-leader-story-arial10.docx');
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  expect(hash(source)).toBe(row.sourceHash);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: `Font range ${row.name}.docx`,
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const story = page.locator(`.word-page-story[data-word-story*="${row.kind}"]`).first();
  const snapshot = (element: Locator) => element.locator('p').evaluateAll(paragraphs => paragraphs.map(p => {
    const read = (text: string, element: Element) => {
      const style = getComputedStyle(element);
      return { text, family: style.fontFamily.replace(/^['"]|['"]$/g, ''), size: Number((parseFloat(style.fontSize) * .75).toFixed(4)) };
    };
    const characters = [], walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const node = walker.currentNode as Text;
      if (!node.data || node.parentElement!.closest('.ProseMirror-widget,[contenteditable=false]')) continue;
      for (const character of node.data) if (character !== '\t') characters.push(read(character, node.parentElement!));
    }
    characters.push(read('\r', p));
    return characters;
  }));
  await expect(page.locator('.section-page')).toHaveCount(2);
  const baseline = await snapshot(story);
  const expected = row.paragraphs.map(p => p.characters.filter(c => c.text !== '\t'));
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${row.kind}`) }).click();
  const dialog = page.getByRole('dialog').last();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await editor.focus(); await page.keyboard.press('Control+Home');
  for (let i = 0; i < row.selection.start; i++) await page.keyboard.press('ArrowRight');
  for (let i = row.selection.start; i < row.selection.end; i++) await page.keyboard.press('Shift+ArrowRight');
  await dialog.getByRole('combobox', { name: 'Font family', exact: true }).selectOption('Times New Roman');
  await dialog.getByRole('spinbutton', { name: 'Font size', exact: true }).fill('20'); await page.keyboard.press('Tab');
  await expect.poll(() => snapshot(editor)).toEqual(expected);
  for (let i = 0; i < 2; i++) await dialog.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(() => snapshot(editor)).toEqual(baseline);
  for (let i = 0; i < 2; i++) await dialog.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect.poll(() => snapshot(editor)).toEqual(expected);
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.insertText('X');
  await expect(editor.locator('p').first()).toHaveText('XA\tB');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText('A\tB');
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText('XA\tB');
  await page.keyboard.press('Control+z'); await expect.poll(() => snapshot(editor)).toEqual(expected);
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await expect.poll(() => snapshot(story)).toEqual(expected);
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); await expect.poll(() => snapshot(story)).toEqual(baseline);
  await page.getByRole('button', { name: 'Redo', exact: true }).click(); await expect.poll(() => snapshot(story)).toEqual(expected);
  const outputs: Record<string, string> = {};
  const download = async (filename: string, label: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click(); await (await pending).saveAs(root + '/' + filename);
    outputs[filename] = hash(await fs.readFile(root + '/' + filename));
  };
  await download('edited.docx', 'DOCX file Editable in Microsoft Word'); await download('edited.pdf', 'PDF file');
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `Font range ${row.name}`, exact: true }).click();
  await expect.poll(() => snapshot(story)).toEqual(expected); await download('reloaded.pdf', 'PDF file');
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles({ name: row.name + '-export.docx',
    buffer: await fs.readFile(root + '/edited.docx'), mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect.poll(() => snapshot(story)).toEqual(expected); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ outputs, errors, sourceHash: hash(source),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-story-font-ranges.json')), buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-story-font-ranges.spec.ts')) }, null, 2));
});
