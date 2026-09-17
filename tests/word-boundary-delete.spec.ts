import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import cases from './fixtures/word-boundary-delete/browser-cases.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
for (const sample of cases.rows) test(`Selected Word boundary ${sample.name}: both keys, history and actual files`, async ({ page }) => {
  const run = process.env.NOFFICE_BOUNDARY_DELETE_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid boundary evidence run');
  const root = `.local/word-boundary-delete/${run}/${sample.name}`;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile(sample.fixture); expect(hash(source)).toBe(sample.sourceHash);
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: sample.name + '.docx', buffer: source,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await waitWordLayout(editor);
  const model = () => editor.evaluate(el => (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getJSON());
  const verify = async () => {
    const state = await editor.evaluate((el, position) => {
      const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
      const node = e.state.doc.nodeAt(position), font = node?.marks.find(m => m.type.name === 'textStyle');
      return { text: e.state.doc.textBetween(0, e.state.doc.content.size, '\r') + '\r',
        paragraphs: e.state.doc.childCount, family: font?.attrs.fontFamily?.replace(/^"(.*)"$/, '$1'),
        size: font?.attrs.fontSize, bold: !!node?.marks.some(m => m.type.name === 'bold'),
        italic: !!node?.marks.some(m => m.type.name === 'italic'),
        markFamily: e.state.doc.firstChild!.attrs.paragraphFontFamily?.replace(/^"(.*)"$/, '$1'),
        markSize: e.state.doc.firstChild!.attrs.paragraphFontSize };
    }, sample.expectedCaret);
    expect(state.text).toBe(sample.expectedText); expect(state.paragraphs).toBe(2);
    if (sample.expectedSpaceFont) expect(state).toMatchObject({ family: sample.expectedSpaceFont.family,
      size: sample.expectedSpaceFont.size + 'pt', bold: !!sample.expectedSpaceFont.bold, italic: !!sample.expectedSpaceFont.italic });
    if (sample.expectedParagraphFont) expect(state).toMatchObject({ markFamily: sample.expectedParagraphFont.markFamily,
      markSize: sample.expectedParagraphFont.markSize + 'pt' });
  };
  const before = await model();
  const keyStates = [];
  for (const key of ['Backspace', 'Delete']) {
    await editor.evaluate(el => {
      const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
      const from = e.state.doc.firstChild!.nodeSize - 1;
      e.commands.setTextSelection({ from, to: from + 2 }); e.view.focus();
    });
    await page.keyboard.press(key); await waitWordLayout(editor); await verify();
    const caret = await editor.evaluate(el => {
      const s = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.state.selection;
      return { from: s.from, to: s.to };
    });
    expect(caret).toEqual({ from: sample.expectedCaret, to: sample.expectedCaret });
    const after = await model(); keyStates.push({ key, after, caret });
    await editor.focus(); await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(before);
    await page.keyboard.press('Control+y'); await expect.poll(model).toEqual(after);
    if (key === 'Backspace') { await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(before); }
  }
  await waitWordLayout(editor);
  const hashes: Record<string, string> = {};
  const capture = async (stage: string) => {
    for (const [extension, label] of [['docx', 'DOCX file Editable in Microsoft Word'], ['pdf', 'PDF file']]) {
      await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: label, exact: true }).click(); const path = `${root}/${stage}.${extension}`;
      await (await pending).saveAs(path); hashes[stage + '.' + extension] = hash(await fs.readFile(path));
    }
  };
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible(); await capture('edited');
  const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click(); await waitWordLayout(editor); await verify(); await capture('reloaded');
  await page.locator('input[type=file][multiple]').setInputFiles(root + '/edited.docx'); await waitWordLayout(editor); await verify();
  await capture('reimported'); expect(hashes['reimported.docx']).toBe(hashes['edited.docx']);
  await fs.writeFile(root + '/report.json', JSON.stringify({ name: sample.name, sourceHash: sample.sourceHash,
    before, keyStates, hashes, buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-boundary-delete.spec.ts')),
    casesHash: hash(await fs.readFile('tests/fixtures/word-boundary-delete/browser-cases.json')),
    nativeReferenceHash: sample.nativeReportHash,
    scope: 'Actual Backspace/Delete on a model-selected boundary, full-model undo/redo, saved reload and three actual DOCX/PDF stages/reimport. Independent native file comparison, pointer selection and native physical keys remain separate.' }, null, 2));
});
