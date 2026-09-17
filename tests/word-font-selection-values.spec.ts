import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-font-selection-values.json' with { type: 'json' };
import edits from './fixtures/native-word-font-selection-edits.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const source of reference.sources) test(`Word font controls ${source.name}: native selection values, edits, history and exports`, async ({ page }) => {
  const run = process.env.NOFFICE_FONT_SELECTION_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
  const root = `.local/word-font-selection-values/${run}/${source.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const input = await fs.readFile(`tests/fixtures/word-font-selection-${source.name.toLowerCase()}.docx`);
  expect(hash(input)).toBe(source.docxHash);
  const expected = edits.rows.find(row => row.name === source.name)!;
  const body = source.kind === 'Body', kind = source.kind === 'Headers' ? 'header' : 'footer';
  const errors: string[] = [];page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: `Font controls ${source.name}.docx`, buffer: input,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const openStory = async () => {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${kind}`) }).click();
  };
  if (!body) await openStory();
  const editor = page.getByRole('textbox', { name: body ? 'Document text' : 'Header or footer text', exact: true });
  const controls = body ? page : page.getByRole('dialog').last();
  const size = controls.getByRole('spinbutton', { name: 'Font size', exact: true });
  const family = controls.getByRole('combobox', { name: 'Font family', exact: true });
  const model = () => editor.evaluate(element => (element as HTMLElement & { editor: {getJSON(): unknown} }).editor.getJSON());
  const characters = () => editor.evaluate(element => {
    type Node = {type: string; text?: string; attrs?: Record<string, unknown>; content?: Node[]; marks?: {type: string; attrs: Record<string, unknown>}[]};
    const p = (element as HTMLElement & {editor: {getJSON(): Node}}).editor.getJSON().content![0];
    const read = (text: string, attrs: Record<string, unknown> = {}) => ({text,
      family: String(attrs.fontFamily || p.attrs!.paragraphFontFamily).replace(/^["']|["']$/g, ''),
      size: parseFloat(String(attrs.fontSize || p.attrs!.paragraphFontSize))});
    const result = (p.content || []).flatMap(n => [...(n.type === 'wordTab' ? '\t' : n.text || '')]
      .map(c => read(c, n.marks?.find(m => m.type === 'textStyle')?.attrs)));
    result.push(read('\r'));return result;
  });
  await expect.poll(characters).toEqual(source.characters);const baseline = await model();
  const select = async (from: number, to: number) => {
    await editor.focus();await page.keyboard.press('Control+Home');
    for (let i = 0; i < from; i++) await page.keyboard.press('ArrowRight');
    for (let i = from; i < to; i++) await page.keyboard.press('Shift+ArrowRight');
  };
  const observed = [];
  for (const row of reference.rows.filter(row => row.source === source.name)) {
    await select(row.from, row.to);
    await expect(size).toHaveValue(row.size === 9999999 ? '' : String(row.size));
    await expect(family).toHaveValue(row.family);
    observed.push({name:row.name,size:await size.inputValue(),family:await family.inputValue()});
  }
  expect(await model()).toEqual(baseline);
  await select(0, source.profile === 'empty' ? 0 : 3);
  const previousSize = await size.inputValue();await size.fill('0');await size.press('Tab');
  await expect(size).toHaveValue(previousSize);expect(await model()).toEqual(baseline);
  await select(0, source.profile === 'empty' ? 0 : 3);
  await size.fill('20');await size.press('Enter');await expect(size).toHaveValue('20');
  const sized = await model();await family.selectOption('Arial');await expect(family).toHaveValue('Arial');
  await expect.poll(characters).toEqual(expected.characters);const formatted = await model();
  const familyChanged = JSON.stringify(sized) !== JSON.stringify(formatted);
  if (familyChanged) {await controls.getByRole('button', {name:'Undo',exact:true}).click();expect(await model()).toEqual(sized);}
  await controls.getByRole('button', {name:'Undo',exact:true}).click();expect(await model()).toEqual(baseline);
  await controls.getByRole('button', {name:'Redo',exact:true}).click();expect(await model()).toEqual(sized);
  if (familyChanged) await controls.getByRole('button', {name:'Redo',exact:true}).click();
  await expect.poll(characters).toEqual(expected.characters);
  if (!body) await page.getByRole('button', {name:'Apply header or footer',exact:true}).click();
  const outputs: Record<string,string> = {};
  const download = async (name: string, label: string) => {
    await page.getByRole('button', {name:'Export',exact:true}).click();const pending=page.waitForEvent('download');
    await page.getByRole('button', {name:label,exact:true}).click();await (await pending).saveAs(root+'/'+name);
    outputs[name]=hash(await fs.readFile(root+'/'+name));
  };
  await download('edited.docx','DOCX file Editable in Microsoft Word');await download('edited.pdf','PDF file');
  await page.reload();await page.getByRole('button', {name:'Recent files',exact:true}).click();
  await page.getByRole('button', {name:`Font controls ${source.name}`,exact:true}).click();
  await download('reloaded.pdf','PDF file');
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles(root+'/edited.docx');
  if (!body) await openStory();await expect.poll(characters).toEqual(expected.characters);expect(errors).toEqual([]);
  await fs.writeFile(root+'/browser-report.json',JSON.stringify({outputs,observed,errors,sourceHash:hash(input),buildHash:await wordStoryBuildHash(),
    testHash:hash(await fs.readFile('tests/word-font-selection-values.spec.ts')),referenceHash:hash(await fs.readFile('tests/fixtures/native-word-font-selection-values.json')),
    editsHash:hash(await fs.readFile('tests/fixtures/native-word-font-selection-edits.json'))},null,2));
});

