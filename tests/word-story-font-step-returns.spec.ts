import { test, expect, type Locator } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-story-font-steps.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word mixed story font step ${row.name}: native re-edit returns through history and reload`, async ({ page }) => {
  test.skip(process.env.NOFFICE_STORY_FONT_STEP_RETURN !== '1', 'Requires independently edited installed Word output');
  const run = process.env.NOFFICE_STORY_FONT_STEP_RETURN_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid font range evidence folder');
  const sourceRun = process.env.NOFFICE_STORY_FONT_STEP_SOURCE_RUN || 'browser-v2';
  if (!/^browser-v\d+$/.test(sourceRun)) throw Error('Invalid native source evidence folder');
  const nativeRoot = `.local/word-story-font-steps/${sourceRun}/${row.name}/native-v1`;
  const root = nativeRoot + "/" + run;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile(nativeRoot + '/returned.docx');
  const nativeBytes = await fs.readFile(nativeRoot + '/native-report.json');
  const native = JSON.parse(nativeBytes.toString().replace(/^\uFEFF/, ''));
  const returned = native.stages.find((stage: { stage: string }) => stage.stage === 'returned');
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  expect(hash(source)).toBe(returned.docxHash);
  expect(hash(await fs.readFile(nativeRoot + '/returned.pdf'))).toBe(returned.pdfHash);
  expect(native.executableHash).toBe(reference.executableHash);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: `Font step return ${row.name}.docx`,
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
  const expected = returned.paragraphs.map((p: { characters: { text: string; family: string; size: number }[] }) => p.characters.filter(c => c.text !== '\t'));
  await expect.poll(() => snapshot(story)).toEqual(expected);
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${row.kind}`) }).click();
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
  await expect.poll(model).toEqual(returned.paragraphs.map((p: { characters: unknown[] }) => p.characters));
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.insertText('X');
  await expect(editor.locator('p').first()).toHaveText('XN\tB');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText('N\tB');
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText('XN\tB');
  await page.keyboard.press('Control+z'); await expect.poll(() => snapshot(editor)).toEqual(expected);
  await expect.poll(model).toEqual(returned.paragraphs.map((p: { characters: unknown[] }) => p.characters));
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await expect.poll(() => snapshot(story)).toEqual(expected);
  const outputs: Record<string, string> = {};
  const download = async (filename: string, label: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click(); await (await pending).saveAs(root + '/' + filename);
    outputs[filename] = hash(await fs.readFile(root + '/' + filename));
  };
  await download('returned.docx', 'DOCX file Editable in Microsoft Word'); await download('returned.pdf', 'PDF file');
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `Font step return ${row.name}`, exact: true }).click();
  await expect.poll(() => snapshot(story)).toEqual(expected); await download('reloaded.pdf', 'PDF file');
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles({ name: row.name + '-export.docx',
    buffer: await fs.readFile(root + '/returned.docx'), mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect.poll(() => snapshot(story)).toEqual(expected); expect(errors).toEqual([]);
  expect(await fs.readFile(root + '/returned.docx')).toEqual(source);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ outputs, errors, sourceHash: hash(source), nativeReceiptHash: hash(nativeBytes),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-story-font-steps.json')), buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-story-font-step-returns.spec.ts')) }, null, 2));
});
