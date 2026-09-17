import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-font-step-carets.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word font caret ${row.name}: caret formatting, typing and native history remain consistent`, async ({ page }) => {
  const run = process.env.NOFFICE_FONT_STEP_CARET_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
  const root = `.local/word-story-font-step-carets/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile(`tests/fixtures/word-font-step-caret-${row.kind}s-${row.content}.docx`);
  expect(hash(source)).toBe(row.sourceHash);
  const expected = row.paragraphs;
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles({ name: `Caret steps ${row.name}.docx`,
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect(page.locator('.section-page')).toHaveCount(2);
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  const open = () => page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${row.kind}`) }).click();
  await open(); const dialog = page.getByRole('dialog').last();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
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
  await editor.focus();await page.keyboard.press('Control+Home');
  for (let i = 0; i < row.offset; i++) await page.keyboard.press('ArrowRight');
  await dialog.getByRole('button', { name: row.mode + ' font', exact: true }).click();
  await expect.poll(model).toEqual(row.formatted.map(p => p.characters));
  await editor.focus();await page.keyboard.insertText('X');
  await expect.poll(model).toEqual(row.typed.map(p => p.characters));
  await page.keyboard.press('Control+z');await expect.poll(model).toEqual(row.undo.map(p => p.characters));
  await page.keyboard.press('Control+y');await expect.poll(model).toEqual(row.redo.map(p => p.characters));
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  const story = page.locator(`.word-page-story[data-word-story*="${row.kind}"]`).first();
  await expect(story.locator('p').first()).toHaveText(expected[0].text.replace(/\r$/, ''));
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  const text = () => story.locator('p').first().textContent();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(text).toBe(baseline[0].map(c => c.text).join('').replace(/\r$/, ''));
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect.poll(text).toBe(expected[0].text.replace(/\r$/, ''));
  const outputs: Record<string,string> = {};
  const download = async (name: string, label: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();await (await pending).saveAs(root+'/'+name);
    outputs[name] = hash(await fs.readFile(root+'/'+name));
  };
  await download('edited.docx','DOCX file Editable in Microsoft Word');await download('edited.pdf','PDF file');
  await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `Caret steps ${row.name}`, exact: true }).click();
  await expect(story.locator('p').first()).toHaveText(expected[0].text.replace(/\r$/, ''));await download('reloaded.pdf','PDF file');
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles(root+'/edited.docx');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();await open();
  await expect.poll(model).toEqual(expected.map(p => p.characters));expect(errors).toEqual([]);
  await fs.writeFile(root+'/browser-report.json',JSON.stringify({outputs,errors,sourceHash:hash(source),buildHash:await wordStoryBuildHash(),
    testHash:hash(await fs.readFile('tests/word-story-font-step-carets.spec.ts')),referenceHash:hash(await fs.readFile('tests/fixtures/native-word-font-step-carets.json'))},null,2));
});
