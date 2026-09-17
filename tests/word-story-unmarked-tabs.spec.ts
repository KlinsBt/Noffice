import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-unmarked-story-tabs.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word ${row.kind}: tab entered before text retains unselected font`, async ({ page }) => {
  const run = process.env.NOFFICE_UNMARKED_TAB_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
  const root = `.local/word-story-unmarked-tabs/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile(`tests/fixtures/word-unmarked-tab-${row.kind}.docx`);
  expect(hash(source)).toBe(row.stages.find(stage => stage.stage === 'source')!.docxHash);
  const expected = row.stages.find(stage => stage.stage === 'expected')!.paragraphs;
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles({ name: `Unmarked ${row.kind}.docx`,
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect(page.locator('.section-page')).toHaveCount(2);
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  const open = () => page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${row.kind}`) }).click();
  await open(); const dialog = page.getByRole('dialog').last();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await editor.focus(); await page.keyboard.press('Control+Home');
  await expect(editor.locator('p').first()).toHaveText('');
  await page.keyboard.press('Tab'); await expect(editor.locator('p').first()).toHaveText('\t');
  await page.keyboard.press('Home'); await page.keyboard.insertText('A');
  await page.keyboard.press('End'); await page.keyboard.insertText('B');
  await expect(editor.locator('p').first()).toHaveText('A\tB');
  const model = () => editor.evaluate(element => {
    type Node = { type: string; text?: string; attrs?: Record<string, unknown>; marks?: {type: string; attrs: Record<string, unknown>}[]; content?: Node[] };
    const editor = (element as HTMLElement & { editor: { getJSON(): Node } }).editor;
    return editor.getJSON().content!.map(p => {
      const read = (text: string, attrs: Record<string, unknown> = {}) => ({text,
        family: String(attrs.fontFamily || p.attrs!.paragraphFontFamily).replace(/^['"]|['"]$/g, ''),
        size: parseFloat(String(attrs.fontSize || p.attrs!.paragraphFontSize)),
      });
      const characters = (p.content || []).flatMap(node => [...(node.type === 'wordTab' ? '\t' : node.text || '')]
        .map(text => read(text, node.marks?.find(mark => mark.type === 'textStyle')?.attrs)));
      characters.push(read('\r'));return characters;
    });
  });
  const baseline = await model();
  await editor.focus(); await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowRight'); await page.keyboard.press('ArrowRight');
  for (let i = 0; i < 3; i++) await page.keyboard.press('Shift+ArrowRight');
  await dialog.getByRole('combobox', { name: 'Font family', exact: true }).selectOption('Times New Roman');
  await dialog.getByRole('spinbutton', { name: 'Font size', exact: true }).fill('20'); await page.keyboard.press('Tab');
  await expect.poll(model).toEqual(expected.map(p => p.characters));
  for (let i = 0; i < 2; i++) await dialog.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(model).toEqual(baseline);
  for (let i = 0; i < 2; i++) await dialog.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect.poll(model).toEqual(expected.map(p => p.characters));
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  const story = page.locator(`.word-page-story[data-word-story*="${row.kind}"]`).first();
  await expect(story.locator('p').first()).toHaveText('A\tB');
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(story.locator('p').first()).toHaveText('');
  await page.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(story.locator('p').first()).toHaveText('A\tB');
  const outputs: Record<string,string> = {};
  const download = async (name: string, label: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();await (await pending).saveAs(root+'/'+name);
    outputs[name] = hash(await fs.readFile(root+'/'+name));
  };
  await download('edited.docx','DOCX file Editable in Microsoft Word');await download('edited.pdf','PDF file');
  await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `Unmarked ${row.kind}`, exact: true }).click();
  await expect(story.locator('p').first()).toHaveText('A\tB');await download('reloaded.pdf','PDF file');
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles(root+'/edited.docx');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();await open();
  await expect.poll(model).toEqual(expected.map(p => p.characters));expect(errors).toEqual([]);
  await fs.writeFile(root+'/browser-report.json',JSON.stringify({outputs,errors,sourceHash:hash(source),buildHash:await wordStoryBuildHash(),
    testHash:hash(await fs.readFile('tests/word-story-unmarked-tabs.spec.ts')),referenceHash:hash(await fs.readFile('tests/fixtures/native-word-unmarked-story-tabs.json'))},null,2));
});
