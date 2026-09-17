import { expect, it } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { wordSelectionFont } from './word-selection-font';
import native from '../tests/fixtures/native-word-font-selection-values.json';

it.each(native.rows)('reflects native aggregate formatting for $name', row => {
  const source = native.sources.find(source => source.name === row.source)!;
  const mark = source.characters.at(-1)!;
  const text = source.characters.slice(0, -1).map(c => `<span style="font-family:${c.family};font-size:${c.size}pt"${c.text === '\t' ? ' data-word-tab="true"' : ''}>${c.text}</span>`).join('');
  const editor = new Editor({ extensions: wordExtensions(), content:
    `<p style="font-family:${mark.family};font-size:${mark.size}pt">${text}</p><p>Next</p>`, parseOptions: { preserveWhitespace: true } });
  try {
    editor.commands.setTextSelection({ from: 1 + row.from, to: row.to === 4 ? 6 : 1 + row.to });
    const before = editor.getJSON(), selection = editor.state.selection.toJSON();
    expect(wordSelectionFont(editor.state)).toEqual({ family: row.family || null, size: row.size === 9999999 ? null : row.size });
    expect(editor.getJSON()).toEqual(before);expect(editor.state.selection.toJSON()).toEqual(selection);
  } finally {editor.destroy();}
});

