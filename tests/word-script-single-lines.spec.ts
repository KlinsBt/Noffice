import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-script-single-lines.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word Single scripts ${row.name}: line height, history, reload and files`, async ({ page }) => {
  const run = process.env.NOFFICE_SCRIPT_SINGLE_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
  const root = `.local/word-script-single-lines/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile('tests/fixtures/word-script-single-lines/' + row.sourceFile);
  expect(hash(source)).toBe(row.stages[0].docxHash);
  const name = 'Single scripts ' + row.name, body = row.kind === 'Body';
  const kind = row.kind === 'Headers' ? 'header' : 'footer';
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const open = async () => {
    if (body) return;
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: new RegExp(`^Section 1 — Default ${kind}`) }).click();
  };
  const editor = page.getByRole('textbox', { name: body ? 'Document text' : 'Header or footer text', exact: true });
  const characters = () => editor.evaluate(element => {
    type Node = { type: string; text?: string; attrs: Record<string, unknown>; content?: Node[]; marks?: { type: string; attrs?: Record<string, unknown> }[] };
    const doc = (element as HTMLElement & { editor: { getJSON(): Node } }).editor.getJSON();
    return doc.content!.map(p => {
      const read = (text: string, marks: Node['marks'] = [], paragraph = false) => {
        const style = marks.find(m => m.type === 'textStyle')?.attrs;
        return { text, family: String(style?.fontFamily || p.attrs.paragraphFontFamily).replace(/^["']|["']$/g, ''),
          size: parseFloat(String(style?.fontSize || p.attrs.paragraphFontSize)),
          superscript: (paragraph ? p.attrs.paragraphScript === 'superscript' : marks.some(m => m.type === 'superscript')) ? -1 : 0,
          subscript: (paragraph ? p.attrs.paragraphScript === 'subscript' : marks.some(m => m.type === 'subscript')) ? -1 : 0 };
      };
      return [...(p.content || []).flatMap(n => [...(n.text || '')].map(c => read(c, n.marks))), read('\r', [], true)];
    });
  });
  const expected = row.stages[0].paragraphs.map(p => p.characters);
  await open(); await expect.poll(characters).toEqual(expected);
  const measure = () => editor.locator('p').evaluateAll(paragraphs => paragraphs.map(p => {
    const style = getComputedStyle(p), scale = p.getBoundingClientRect().width / parseFloat(style.width);
    return { height: p.getBoundingClientRect().height / scale * .75, line: parseFloat(style.lineHeight) * .75 };
  }));
  // Mixed-leading wrappers deliberately use zero CSS leading and derive the
  // visible paragraph advance from their run struts. Compare the actual box.
  await expect.poll(async () => (await measure()).every(p => Math.abs(p.height - 11.5) <= .15)).toBe(true);
  const lineMetrics = await measure();
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.insertText('Q');
  await expect(editor.locator('p').first()).toHaveText('QAb');
  await page.keyboard.press('Control+z'); await expect.poll(characters).toEqual(expected);
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText('QAb');
  await page.keyboard.press('Control+z'); await expect.poll(characters).toEqual(expected);
  if (!body) await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const outputs: Record<string, string> = {};
  const download = async (file: string, label: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await pending).saveAs(root + '/' + file);
    outputs[file] = hash(await fs.readFile(root + '/' + file));
  };
  await download('recovered.docx', 'DOCX file Editable in Microsoft Word');
  await download('recovered.pdf', 'PDF file');
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click();
  await open(); await expect.poll(characters).toEqual(expected);
  if (!body) await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  await download('reloaded.pdf', 'PDF file');
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/recovered.docx');
  await open(); await expect.poll(characters).toEqual(expected);
  if (!body) await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  await download('reimported.pdf', 'PDF file');
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ outputs, errors, lineMetrics,
    sourceHash: hash(source), buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-script-single-lines.spec.ts')),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-script-single-lines.json')) }, null, 2));
});
