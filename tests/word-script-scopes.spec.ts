import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import native from './fixtures/native-word-script-scopes.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of native.rows) test(`Word script scopes ${row.name}: real editing, history, reload and files`, async ({ page }) => {
  test.setTimeout(60000);
  const run = process.env.NOFFICE_SCRIPT_SCOPES_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
  const root = `.local/word-script-scopes/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (value: Buffer) => createHash('sha256').update(value).digest('hex');
  const source = native.sources.find(s => s.name === row.source)!;
  const bytes = await fs.readFile('tests/fixtures/word-script-scopes/' + source.file);
  expect(hash(bytes)).toBe(source.docxHash);
  const body = row.kind === 'Body', kind = row.kind === 'Headers' ? 'header' : 'footer';
  const name = 'Script scope ' + row.name, errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: bytes,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const open = async () => {
    if (body) return;
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: new RegExp(`^Section 1 — Default ${kind}`) }).click();
  };
  const editor = page.getByRole('textbox', { name: body ? 'Document text' : 'Header or footer text', exact: true });
  const controls = body ? page : page.getByRole('dialog').last();
  const snapshot = () => editor.evaluate(element => {
    type Node = { type: string; text?: string; attrs: Record<string, unknown>; content?: Node[];
      marks?: { type: string; attrs: Record<string, unknown> }[] };
    const p = (element as HTMLElement & { editor: { getJSON(): Node } }).editor.getJSON().content![0];
    const read = (text: string, marks: Node['marks'] = [], paragraph = false) => {
      const attrs = marks?.find(m => m.type === 'textStyle')?.attrs;
      return { text, family: String(attrs?.fontFamily || p.attrs.paragraphFontFamily).replace(/^["']|["']$/g, ''),
        size: parseFloat(String(attrs?.fontSize || p.attrs.paragraphFontSize)),
        superscript: (paragraph ? p.attrs.paragraphScript === 'superscript' : marks?.some(m => m.type === 'superscript')) ? -1 : 0,
        subscript: (paragraph ? p.attrs.paragraphScript === 'subscript' : marks?.some(m => m.type === 'subscript')) ? -1 : 0 };
    };
    return [...(p.content || []).flatMap(n => [...(n.type === 'wordTab' ? '\t' : n.text || '')].map(c => read(c, n.marks))), read('\r', [], true)];
  });
  const selection = () => editor.evaluate(element => {
    const e = (element as HTMLElement & { editor: { state: { selection: { from: number; to: number } } } }).editor;
    return { from: e.state.selection.from, to: e.state.selection.to };
  });
  const select = async () => {
    await editor.focus(); await page.keyboard.press('Control+Home');
    for (let i = 0; i < row.selection.start; i++) {
      await page.keyboard.press('ArrowRight'); await expect.poll(selection).toEqual({ from: i + 2, to: i + 2 });
    }
    for (let i = row.selection.start; i < row.selection.end; i++) await page.keyboard.press('Shift+ArrowRight');
  };
  const click = async (label: string) => controls.getByRole('button', { name: label, exact: true }).click();
  const apply = async () => { if (!body) await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click(); };
  const outputs: Record<string, string> = {};
  const download = async (file: string) => {
    await fs.writeFile(root + '/' + file + '-model.json', JSON.stringify(await page.getByRole('textbox', { name: 'Document text', exact: true }).evaluate(element => {
      const editor = (element as HTMLElement & { editor: { getJSON(): unknown; getHTML(): string } }).editor;
      return { json: editor.getJSON(), html: editor.getHTML() };
    }), null, 2));
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    pending.catch(() => {});
    await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
    const result = await Promise.race([pending, page.locator('.error-banner').waitFor({ state: 'visible' }).then(async () => {
      throw Error(await page.locator('.error-banner').innerText());
    })]);
    await result.saveAs(root + '/' + file); outputs[file] = hash(await fs.readFile(root + '/' + file));
  };
  await open(); await expect.poll(snapshot).toEqual(source.characters); await select();
  const beforeSelection = await selection();
  await click(row.script); await expect.poll(snapshot).toEqual(row.formatted);
  await expect(controls.getByRole('button', { name: row.script, exact: true })).toHaveAttribute('aria-pressed', 'true');
  await click('Undo'); expect(await snapshot()).toEqual(source.characters);
  await click('Redo'); expect(await snapshot()).toEqual(row.formatted);
  // Native format Undo/Redo restores the mark but clears the transient end-caret
  // script choice. Reapply that choice before the separate typing workflow.
  if (body && row.scope === 'end-caret') {
    await expect(controls.getByRole('button', { name: row.script, exact: true })).toHaveAttribute('aria-pressed', 'false');
    await click(row.script); expect(await snapshot()).toEqual(row.formatted);
  }
  await apply(); await download('formatted.docx');
  if (!body) { await open(); await select(); }
  // Re-establish typing state with the measured command after reopening a story.
  // Saving a paragraph mark persists its format; caret stored marks are transient.
  if (!body && row.selection.start === row.selection.end) {
    if (await controls.getByRole('button', { name: row.script, exact: true }).getAttribute('aria-pressed') !== 'true')
      await click(row.script);
  }
  await editor.focus(); await page.keyboard.insertText('X');
  await expect.poll(snapshot).toEqual(row.typed);
  await page.keyboard.press('Control+z'); expect(await snapshot()).toEqual(row.formatted);
  await page.keyboard.press('Control+y'); expect(await snapshot()).toEqual(row.typed);
  await apply(); await download('typed.docx');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click(); await open(); expect(await snapshot()).toEqual(row.typed);
  if (!body) await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/formatted.docx');
  await open(); expect(await snapshot()).toEqual(row.formatted);
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ outputs, errors, beforeSelection,
    sourceHash: hash(bytes), buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-script-scopes.spec.ts')),
    referenceHash: hash(await fs.readFile('tests/fixtures/native-word-script-scopes.json')) }, null, 2));
});
