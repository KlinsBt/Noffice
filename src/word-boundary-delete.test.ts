import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { Editor } from '@tiptap/core';
import { readDocx } from './docx-import';
import { wordExtensions } from './word-extensions';
import { WordEditorSections, updateWordSectionSource } from './word-editor-sections';
import reference from '../tests/fixtures/word-boundary-delete/reference.json';
import fonts from '../tests/fixtures/word-boundary-delete/fonts.json';

vi.mock('mammoth', async original => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});

for (const sample of reference.rows) it(`matches native selected ${sample.name} deletion, caret and history`, async () => {
  const bytes = Uint8Array.from(readFileSync(sample.fixture)).buffer;
  expect(createHash('sha256').update(new Uint8Array(bytes)).digest('hex')).toBe(sample.sourceHash);
  const { content } = await readDocx(bytes);
  for (const native of sample.states) {
    const editor = new Editor({ extensions: [...wordExtensions(), WordEditorSections.configure({ source: content.docxStructure })],
      content: content.html, parseOptions: { preserveWhitespace: true } });
    try {
      updateWordSectionSource(editor, content.docxStructure);
      const from = editor.state.doc.firstChild!.nodeSize - 1;
      editor.commands.setTextSelection({ from, to: from + 2 });
      const before = editor.getJSON();
      expect(editor.can().deleteWordBoundary(native.smart)).toBe(true);
      expect(editor.getJSON()).toEqual(before);
      expect(editor.commands.deleteWordBoundary(native.smart)).toBe(true);
      const after = editor.getJSON();
      expect(editor.state.doc.textBetween(0, editor.state.doc.content.size, '\r') + '\r').toBe(native.after);
      expect(editor.state.doc.childCount).toBe(native.paragraphs);
      expect(editor.state.selection.from).toBe(native.afterSelection.from + 1);
      expect(editor.state.selection.empty).toBe(true);
      editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
      editor.commands.redo(); expect(editor.getJSON()).toEqual(after);
    } finally { editor.destroy(); }
  }
});

it('keeps raw transforms separate from the native selected-boundary command', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p>Alpha.</p><p>Beta</p>' });
  try {
    editor.view.dispatch(editor.state.tr.delete(7, 9));
    expect(editor.state.doc.textContent).toBe('Alpha.Beta');
    editor.commands.undo(); expect(editor.state.doc.childCount).toBe(2);
  } finally { editor.destroy(); }
});

for (const sample of fonts.rows) it(`retains native automatic-space formatting for ${sample.name}`, async () => {
  const bytes = Uint8Array.from(readFileSync(sample.fixture)).buffer;
  expect(createHash('sha256').update(new Uint8Array(bytes)).digest('hex')).toBe(sample.sourceHash);
  const { content } = await readDocx(bytes);
  const editor = new Editor({ extensions: [...wordExtensions(), WordEditorSections.configure({ source: content.docxStructure })],
    content: content.html, parseOptions: { preserveWhitespace: true } });
  try {
    updateWordSectionSource(editor, content.docxStructure);
    const from = editor.state.doc.firstChild!.nodeSize - 1;
    editor.commands.setTextSelection({ from, to: from + 2 }); const before = editor.getJSON();
    expect(editor.commands.deleteWordBoundary()).toBe(true);
    const native = sample.states.find(s => s.smart)!;
    const expected = native.characters.find(c => c.position === native.from)!;
    expect(expected.text).toBe(' ');
    const space = editor.state.doc.nodeAt(from)!;
    expect(editor.state.doc.textBetween(from, from + 1)).toBe(' ');
    const style = space.marks.find(m => m.type.name === 'textStyle');
    expect(style?.attrs.fontFamily?.replace(/^"(.*)"$/, '$1')).toBe(expected.family);
    expect(style?.attrs.fontSize).toBe(expected.size + 'pt');
    expect(space.marks.some(m => m.type.name === 'bold')).toBe(!!expected.bold);
    expect(space.marks.some(m => m.type.name === 'italic')).toBe(!!expected.italic);
    const after = editor.getJSON(); editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
    editor.commands.redo(); expect(editor.getJSON()).toEqual(after);
  } finally { editor.destroy(); }
});

it('declines collapsed, partial-text and nested selections without changing their model', () => {
  for (const [html, from, to] of [
    ['<p>Alpha.</p><p>Beta</p>', 7, 7],
    ['<p>Alpha.</p><p>Beta</p>', 6, 10],
    ['<blockquote><p>Alpha.</p><p>Beta</p></blockquote>', 8, 10],
  ] as const) {
    const editor = new Editor({ extensions: wordExtensions(), content: html });
    try {
      editor.commands.setTextSelection({ from, to }); const before = editor.getJSON();
      expect(editor.commands.deleteWordBoundary()).toBe(false);
      expect(editor.getJSON()).toEqual(before);
    } finally { editor.destroy(); }
  }
});
