import { expect, it } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';

function type(editor: Editor, text: string) {
  const { from, to } = editor.state.selection;
  return editor.view.someProp('handleTextInput', handler =>
    handler(editor.view, from, to, text, () => editor.state.tr.insertText(text)));
}
const hyphens = (editor: Editor) => {
  const nodes: { kind: string; marks: unknown[] }[] = [];
  editor.state.doc.descendants(node => {
    if (node.type.name === 'wordHyphen') nodes.push({ kind: node.attrs.kind, marks: node.marks.map(m => m.toJSON()) });
  });
  return nodes;
};

it('retains formatting, selection and atomic history when committed text contains multiple literal hyphens', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p><strong><em>ABCD</em></strong></p>' });
  try {
    editor.commands.setTextSelection({ from: 2, to: 4 }); const before = editor.getJSON();
    expect(type(editor, 'X\u00ad\u00adY')).toBe(true);
    expect(hyphens(editor)).toHaveLength(2);
    for (const node of hyphens(editor)) {
      expect(node.kind).toBe('literal');
      expect(node.marks).toEqual(expect.arrayContaining([{ type: 'bold' }, { type: 'italic' }]));
    }
    expect(editor.state.selection.from).toBe(6); expect(editor.state.selection.empty).toBe(true);
    const after = editor.getJSON(); editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
    editor.commands.redo(); expect(editor.getJSON()).toEqual(after);
    const reloaded = new Editor({ extensions: wordExtensions(), content: editor.getHTML() });
    try { expect(reloaded.getJSON()).toEqual(after); } finally { reloaded.destroy(); }
  } finally { editor.destroy(); }
});

for (const kind of ['literal', 'optional'] as const) it(`materializes an empty paragraph's font for ${kind} input`, () => {
  const editor = new Editor({ extensions: wordExtensions(),
    content: '<p data-word-paragraph-kerning="24" style="font-size:11pt;font-family:Calibri"></p>' });
  try {
    const before = editor.getJSON();
    if (kind === 'literal') expect(type(editor, '\u00ad')).toBe(true);
    else expect(editor.commands.insertWordHyphen('optional')).toBe(true);
    expect(hyphens(editor)).toHaveLength(1);
    const node = editor.state.doc.firstChild!.firstChild!;
    expect(node.type.name).toBe('wordHyphen'); expect(node.attrs.kind).toBe(kind);
    const font = node.marks.find(m => m.type.name === 'textStyle')!;
    expect(font?.attrs).toMatchObject({ fontSize: '11pt', fontFamily: 'Calibri', wordKerning: 24 });
    const after = editor.getJSON(); editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
    editor.commands.redo(); expect(editor.getJSON()).toEqual(after);
  } finally { editor.destroy(); }
});

it('keeps the measured paragraph-mark replacement rule when literal input replaces the last selected text', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p>AB</p><p>C</p>' });
  try {
    editor.commands.setTextSelection({ from: 2, to: 5 }); const before = editor.getJSON();
    expect(type(editor, '\u00adQ')).toBe(true);
    expect(editor.state.doc.childCount).toBe(2); expect(editor.state.doc.child(1).textContent).toBe('C');
    expect(hyphens(editor).map(n => n.kind)).toEqual(['literal']);
    expect(editor.state.doc.child(0).textBetween(0, editor.state.doc.child(0).content.size, '', '-')).toBe('A-Q');
    const after = editor.getJSON(); editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
    editor.commands.redo(); expect(editor.getJSON()).toEqual(after);
  } finally { editor.destroy(); }
});

it('leaves ordinary text and active composition to their existing handlers', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p>AB</p>' });
  try {
    const before = editor.getJSON(); expect(type(editor, 'normal')).not.toBe(true);
    Object.defineProperty(editor.view, 'composing', { configurable: true, get: () => true });
    expect(type(editor, '\u00ad')).not.toBe(true); expect(editor.getJSON()).toEqual(before);
  } finally { editor.destroy(); }
});

it('converts only the inserted literal characters and retains unrelated existing controls', () => {
  const editor = new Editor({ extensions: wordExtensions(),
    content: '<p>raw\u00ad<span data-word-hyphen="optional">\u00ad</span>AB</p>' });
  try {
    editor.commands.setTextSelection(editor.state.doc.firstChild!.nodeSize - 1);
    expect(type(editor, 'x\u00ady')).toBe(true);
    expect(hyphens(editor).map(n => n.kind)).toEqual(['optional', 'literal']);
    expect(editor.state.doc.firstChild!.firstChild!.text).toBe('raw\u00ad');
  } finally { editor.destroy(); }
});
