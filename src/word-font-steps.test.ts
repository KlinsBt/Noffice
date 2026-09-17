import { expect, it } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import native from '../tests/fixtures/native-word-story-font-steps.json';

const paragraph = '<p style="font-family:Arial;font-size:10pt">A<span data-word-tab="true">\t</span>B</p>';
const initial = paragraph.replace('B</p>', '<span style="font-size:20pt">B</span></p>') + paragraph.repeat(4);
function snapshot(editor: Editor) {
  return editor.getJSON().content!.map(p => {
    const read = (text: string, attrs: Record<string, unknown> = {}) => ({ text,
      family: String(attrs.fontFamily || p.attrs!.paragraphFontFamily).replace(/^['"]|['"]$/g, ''),
      size: parseFloat(String(attrs.fontSize || p.attrs!.paragraphFontSize)),
    });
    const chars = p.content!.flatMap(node => {
      const text = node.type === 'wordTab' ? '\t' : 'text' in node ? node.text : '';
      expect(text).not.toBe('');
      return [...text].map(char => read(char, node.marks?.find(mark => mark.type === 'textStyle')?.attrs));
    });
    chars.push(read('\r'));return chars;
  });
}

it.each(native.rows)('matches native mixed-size selection $name without collapsing its runs', row => {
  const editor = new Editor({ extensions: wordExtensions(), content: initial });
  try {
    editor.commands.setTextSelection({ from: 1, to: row.scope === 'text' ? 4 : 6 });
    const before = editor.getJSON(), selection = editor.state.selection.toJSON();
    expect(editor.can().stepWordFontSize(row.mode === 'Grow')).toBe(true);
    expect(editor.getJSON()).toEqual(before);
    editor.commands.stepWordFontSize(row.mode === 'Grow');
    expect(snapshot(editor)).toEqual(row.paragraphs.map(p => p.characters));
    expect(editor.state.selection.toJSON()).toEqual(selection);
    const after = editor.getJSON();
    editor.commands.undo();expect(editor.getJSON()).toEqual(before);
    editor.commands.redo();expect(editor.getJSON()).toEqual(after);
    const reloaded = new Editor({ extensions: wordExtensions(), content: editor.getHTML() });
    try {expect(snapshot(reloaded)).toEqual(snapshot(editor));} finally {reloaded.destroy();}
  } finally {editor.destroy();}
});

it('rejects a malformed mixed run atomically, including can() checks', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: initial });
  try {
    editor.chain().setTextSelection({ from: 3, to: 4 }).setMark('textStyle', { fontSize: 'NaNpt' }).run();
    editor.commands.setTextSelection({ from: 1, to: 4 });
    const before = editor.getJSON();
    expect(editor.can().stepWordFontSize(true)).toBe(false);
    expect(editor.commands.stepWordFontSize(true)).toBe(false);
    expect(editor.getJSON()).toEqual(before);
  } finally { editor.destroy(); }
});

it.each([[1638, true], [1, false]] as const)('keeps bound %s unchanged without adding history', (size, grow) => {
  const editor = new Editor({ extensions: wordExtensions(), content: `<p style="font-size:${size}pt">Limit</p>` });
  try {
    editor.commands.selectAll();const before = editor.getJSON();
    expect(editor.can().stepWordFontSize(grow)).toBe(true);
    expect(editor.commands.stepWordFontSize(grow)).toBe(true);
    expect(editor.getJSON()).toEqual(before);expect(editor.can().undo()).toBe(false);
  } finally { editor.destroy(); }
});
