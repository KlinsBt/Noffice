import { it, expect, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { wordExtensions, wordJSON } from './word-extensions';
import { wordSelectionParagraphLayout, wordParagraphIndentInput, type WordParagraphLayout } from './word-paragraph-layout-commands';
import { readDocx } from './docx-import';
import { exportOffice, importFile } from './formats';
import native from '../tests/fixtures/native-word-story-paragraph-layout.json';
import inputs from '../tests/fixtures/native-word-paragraph-indent-input.json';
vi.mock('mammoth', async original => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});
const semantic = (p: typeof native.rows[number]['states'][number]['paragraphs'][number]): WordParagraphLayout => ({
  textAlign: (['left', 'center', 'right', 'justify'] as const)[p.alignment],
  indentStart: p.indentStart, indentEnd: p.indentEnd, firstLineIndent: p.firstLineIndent,
  keepNext: p.keepNext === -1, keepLines: p.keepLines === -1,
  pageBreakBefore: p.pageBreakBefore === -1, widowControl: p.widowControl === -1,
});
for (const row of inputs.rows) it(`matches independently saved Word indent input ${row.name}`, () => {
  expect(wordParagraphIndentInput(row.input)).toBe(row.accepted ? row.savedTwips / 20 : null);
});
for (const row of native.rows) it(`authors ${row.name} with native paragraph semantics and retained DOCX`, async () => {
  const source = row.states.find(s => s.stage === 'source')!;
  const wanted = row.states.find(s => s.stage === 'edited')!;
  const data = new Uint8Array(readFileSync(source.fixture)).buffer;
  const file = await importFile(Object.assign(new File([data], row.name + '.docx'), { arrayBuffer: async () => data }));
  if (file.content.kind !== 'word') throw Error('Word fixture');
  const ref = file.content.docxStructure!.sections[0][row.kind === 'header' ? 'headers' : 'footers'].find(r => r.type === 'default')!;
  const story = file.content.stories!.parts.find(p => p.kind === row.kind && p.relationshipIds.includes(ref.relationshipId))!;
  const editor = new Editor({ extensions: wordExtensions(), content: story.html });
  try {
    let pos = 1; editor.state.doc.forEach((node, offset, i) => { if (i === 2) pos = offset + 1; });
    editor.commands.setTextSelection(pos);
    expect(wordSelectionParagraphLayout(editor.state)).toEqual(semantic(source.paragraphs[2]));
    const before = editor.getJSON(), target = semantic(wanted.paragraphs[2]);
    expect(editor.can().setWordParagraphLayout(target)).toBe(true); expect(editor.can().undo()).toBe(false);
    expect(editor.commands.setWordParagraphLayout(target)).toBe(true);
    expect(wordSelectionParagraphLayout(editor.state)).toEqual(target);
    const after = editor.getJSON(); editor.commands.setWordParagraphLayout(target);
    editor.commands.undo(); expect(editor.getJSON()).toEqual(before); expect(editor.can().undo()).toBe(false);
    editor.commands.redo(); expect(editor.getJSON()).toEqual(after);
    story.html = editor.getHTML();
    const zip = await JSZip.loadAsync(await exportOffice(file));
    const result = (await readDocx(await zip.generateAsync({ type: 'arraybuffer' }))).content;
    const output = wordJSON(result.stories!.parts.find(p => p.path === story.path)!.html).content!;
    expect(output).toHaveLength(5);
    const reopened = new Editor({ extensions: wordExtensions(), content: { type: 'doc', content: output } });
    try {
      reopened.state.doc.forEach((node, offset, i) => {
        reopened.commands.setTextSelection(offset + 1);
        expect(wordSelectionParagraphLayout(reopened.state)).toEqual(semantic(wanted.paragraphs[i]));
      });
    } finally { reopened.destroy(); }
    const original = await JSZip.loadAsync(data);
    for (const [path, entry] of Object.entries(original.files)) if (!entry.dir && path !== story.path)
      expect(await zip.file(path)!.async('uint8array'), path).toEqual(await entry.async('uint8array'));
  } finally { editor.destroy(); }
});
it('reads mixed properties independently and applies a partial patch atomically to the selected paragraphs', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p style="margin-inline-start:12pt;text-align:center;break-after:avoid">A</p><h2 style="margin-inline-start:16px;text-align:right;orphans:1">B</h2><p>C</p>' });
  try {
    editor.commands.setTextSelection({ from: 1, to: 7 });
    const before = editor.getJSON(), selected = wordSelectionParagraphLayout(editor.state);
    expect(selected.indentStart).toBe(12); expect(selected.textAlign).toBeNull(); expect(selected.keepNext).toBeNull();
    expect(selected.widowControl).toBeNull();
    editor.commands.setWordParagraphLayout({ indentStart: 36, textAlign: 'justify', keepLines: true });
    const after = editor.getJSON(); expect(after.content![2]).toEqual(before.content![2]);
    expect(after.content![0].attrs!.keepNext).toBe(true); expect(after.content![1].attrs!.widowControl).toBe(false);
    expect(wordSelectionParagraphLayout(editor.state)).toMatchObject({ indentStart: 36, textAlign: 'justify', keepLines: true });
    editor.commands.setWordParagraphLayout({ indentStart: 36 });
    editor.commands.undo(); expect(editor.getJSON()).toEqual(before); expect(editor.can().undo()).toBe(false);
    editor.commands.redo(); expect(editor.getJSON()).toEqual(after);
  } finally { editor.destroy(); }
});
it('rejects invalid or partly invalid formatting atomically, leaving selection, content and history intact', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p>Safe</p>' });
  try {
    const before = editor.getJSON(), selection = editor.state.selection.toJSON();
    for (const patch of [null, {}, [], { textAlign: 'bad' }, { direction: 'rtl' }, { keepNext: 1 },
      { textAlign: { toString: () => 'left' } },
      ...[-1585, 1585, NaN, Infinity, '36', null].map(indentStart => ({ indentStart, keepLines: true }))]) {
      expect(editor.commands.setWordParagraphLayout(patch as Partial<WordParagraphLayout>)).toBe(false);
      expect(editor.getJSON()).toEqual(before); expect(editor.state.selection.toJSON()).toEqual(selection);
      expect(editor.can().undo()).toBe(false);
    }
    expect(editor.commands.setWordParagraphLayout({ indentStart: 0, textAlign: 'left', widowControl: true })).toBe(true);
    expect(editor.getJSON()).toEqual(before); expect(editor.can().undo()).toBe(false);
  } finally { editor.destroy(); }
});
