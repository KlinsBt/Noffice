import { afterEach, expect, it } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import {
  changeCase,
  nextFontSize,
  captureTextFormat,
  applyTextFormat,
  indentParagraphs,
} from './word-commands';
import { showFormattingMarks } from './word-marks';
const editors: Editor[] = [];
afterEach(() => {
  for (const editor of editors) editor.destroy();
  editors.length = 0;
});
function make(html: string) {
  const editor = new Editor({ extensions: wordExtensions(), content: html });
  editors.push(editor);
  return editor;
}
it.each([
  ['upper', 'HELLO WORLD. SECOND SENTENCE!'],
  ['lower', 'hello world. second sentence!'],
  ['title', 'Hello World. Second Sentence!'],
  ['sentence', 'Hello world. Second sentence!'],
  ['toggle', 'hELLO wORLD. sECOND sENTENCE!'],
] as const)('changes %s case across differently marked text', (mode, want) => {
  const editor = make(
    '<p data-source-paragraph="0:7"><strong>Hello </strong><em>World. Second Sentence!</em></p>',
  );
  editor.commands.setTextSelection({ from: 1, to: editor.state.doc.content.size - 1 });
  changeCase(editor, mode);
  expect(editor.getText()).toBe(want);
  expect(editor.getJSON().content![0].attrs!.sourceParagraph).toBe('0:7');
  expect(editor.getHTML()).toContain('<strong>');
  expect(editor.getHTML()).toContain('<em>');
  editor.commands.undo();
  expect(editor.getText()).toBe('Hello World. Second Sentence!');
});
it('handles expanding case mappings without moving inline objects or corrupting selection', () => {
  const editor = make('<p>A <strong>straße</strong><span data-word-tab="true"></span>z</p>');
  editor.commands.setTextSelection({ from: 3, to: 9 });
  changeCase(editor, 'upper');
  expect(editor.getText()).toContain('STRASSE');
  expect(editor.state.selection.to - editor.state.selection.from).toBe(7);
  expect(editor.getHTML()).toContain('data-word-tab');
  expect(editor.view.dom.querySelector('strong')?.textContent).toBe('STRASSE');
});
it('copies character formatting without copying links or source paragraph identities', () => {
  const editor = make(
    '<p data-source-paragraph="0:1"><strong><span style="font-size:18pt;color:#ff0000">Source</span></strong></p><p data-source-paragraph="0:2"><a href="https://example.com"><em>Target</em></a></p>',
  );
  editor.commands.setTextSelection({ from: 1, to: 7 });
  const format = captureTextFormat(editor);
  editor.commands.setTextSelection({ from: 9, to: 15 });
  applyTextFormat(editor, format);
  const target = editor.getJSON().content![1];
  expect(target.attrs!.sourceParagraph).toBe('0:2');
  const marks = target.content![0].marks!;
  expect(marks.map((m) => m.type)).toContain('bold');
  expect(marks.map((m) => m.type)).not.toContain('italic');
  expect(marks.map((m) => m.type)).toContain('link');
  editor.commands.undo();
  expect(editor.getHTML()).toContain('<em>Target</em>');
});
it('changes paragraph indents as separate undoable operations and retains mappings', () => {
  const editor = make(
    '<p data-source-paragraph="0:1" style="margin-inline-start:12pt">First</p><p data-source-paragraph="0:2">Second</p>',
  );
  editor.commands.selectAll();
  indentParagraphs(editor);
  expect(editor.getJSON().content!.map((n) => n.attrs!.indentStart)).toEqual(['48pt', '36pt']);
  indentParagraphs(editor, true);
  expect(editor.getJSON().content!.map((n) => n.attrs!.indentStart)).toEqual(['12pt', '0pt']);
  editor.commands.undo();
  expect(editor.getJSON().content![0].attrs!.indentStart).toBe('48pt');
  expect(editor.getJSON().content![1].attrs!.sourceParagraph).toBe('0:2');
});
it('uses list nesting commands inside list items', () => {
  const editor = make('<ul><li><p>First</p></li><li><p>Second</p></li></ul>');
  let target = 0;
  editor.state.doc.descendants((node, pos) => {
    if (node.isText && node.text === 'Second') target = pos;
  });
  editor.commands.setTextSelection(target);
  expect(indentParagraphs(editor)).toBe(true);
  expect(editor.getHTML()).toContain('<p>First</p><ul>');
  expect(indentParagraphs(editor, true)).toBe(true);
  expect(editor.getHTML().match(/<ul>/g)).toHaveLength(1);
});
it('renders formatting decorations without inserting characters into saved content', () => {
  const editor = make('<p>Two words<span data-word-tab="true"></span>Tab<br>Line</p>');
  const before = editor.getHTML();
  showFormattingMarks(editor, true);
  expect(editor.view.dom.querySelectorAll('.word-format-mark')).toHaveLength(3);
  expect(editor.view.dom.querySelectorAll('.word-space-mark')).toHaveLength(1);
  expect(editor.getHTML()).toBe(before);
  showFormattingMarks(editor, false);
  expect(editor.view.dom.querySelectorAll('.word-format-mark')).toHaveLength(0);
  expect(editor.getHTML()).toBe(before);
});
it('steps font size through the supported scale and clamps its bounds', () => {
  expect(nextFontSize(12, true)).toBe(14);
  expect(nextFontSize(14, false)).toBe(12);
  expect(nextFontSize(13, true)).toBe(14);
  expect(nextFontSize(13, false)).toBe(12);
  expect(nextFontSize(160, true)).toBe(170);
  expect(nextFontSize(8, false)).toBe(7);
});
// Word 16.0.4627 native Font.Grow/Shrink measurements; see verify-word-font-sizes.ps1.
it.each([
  [1, 2, 1],
  [1.5, 2.5, 1],
  [6, 7, 5],
  [7.5, 8, 6.5],
  [8, 9, 7],
  [9, 10, 8],
  [10, 11, 9],
  [13, 14, 12],
  [28, 36, 26],
  [36, 48, 28],
  [48, 72, 36],
  [71, 72, 48],
  [72, 80, 48],
  [73, 80, 72],
  [79, 80, 72],
  [80, 90, 72],
  [96, 100, 90],
  [160, 170, 150],
  [160.5, 170, 160],
  [200, 210, 190],
  [1638, 1638, 1630],
])('matches native growth and shrinkage at %s points', (size, grow, shrink) => {
  expect(nextFontSize(size, true)).toBe(grow);
  expect(nextFontSize(size, false)).toBe(shrink);
});
