import { expect, it } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import native from '../tests/fixtures/native-word-font-step-carets.json';

const paragraph = '<p style="font-family:Arial;font-size:10pt">A<span data-word-tab="true">\t</span>B</p>';
function snapshot(editor: Editor) {
  return editor.getJSON().content!.map(p => {
    const read = (text: string, attrs: Record<string, unknown> = {}) => ({ text,
      family: String(attrs.fontFamily || p.attrs!.paragraphFontFamily).replace(/^["']|["']$/g, ''),
      size: parseFloat(String(attrs.fontSize || p.attrs!.paragraphFontSize)),
    });
    const chars = (p.content || []).flatMap(node => {
      const text = node.type === 'wordTab' ? '\t' : 'text' in node ? node.text : '';
      return [...text].map(char => read(char, node.marks?.find(mark => mark.type === 'textStyle')?.attrs));
    });
    chars.push(read('\r'));return chars;
  });
}

it.each(native.rows)('matches native caret formatting, typing and undo/redo: $name', row => {
  const first = row.content === 'empty' ? '<p style="font-family:Arial;font-size:10pt"></p>'
    : paragraph.replace('A<span', 'A<span style="font-size:20pt">b</span><span');
  const editor = new Editor({ extensions: wordExtensions(), content: first + paragraph.repeat(4) });
  try {
    editor.commands.setTextSelection(1 + row.offset);
    const before = editor.getJSON(), selection = editor.state.selection.toJSON();
    expect(editor.can().stepWordFontSize(row.mode === 'Grow')).toBe(true);
    expect(editor.getJSON()).toEqual(before);
    expect(editor.commands.stepWordFontSize(row.mode === 'Grow')).toBe(true);
    expect(snapshot(editor)).toEqual(row.formatted.map(p => p.characters));
    expect(editor.state.selection.toJSON()).toEqual(selection);
    editor.view.dispatch(editor.state.tr.insertText('X'));
    expect(snapshot(editor)).toEqual(row.typed.map(p => p.characters));
    editor.commands.undo();expect(snapshot(editor)).toEqual(row.undo.map(p => p.characters));
    editor.commands.redo();expect(snapshot(editor)).toEqual(row.redo.map(p => p.characters));
    const reloaded = new Editor({ extensions: wordExtensions(), content: editor.getHTML() });
    try {expect(snapshot(reloaded)).toEqual(row.paragraphs.map(p => p.characters));} finally {reloaded.destroy();}
  } finally {editor.destroy();}
});
