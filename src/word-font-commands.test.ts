import { expect, it } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import storyRanges from '../tests/fixtures/native-word-story-font-ranges.json';

it('preserves an unselected inherited tab font when a cross-paragraph selection changes the mark', () => {
  const paragraph = '<p style="font-family:Arial;font-size:10pt">A<span data-word-tab="true">\t</span>B</p>';
  const editor = new Editor({ extensions: wordExtensions(), content: paragraph + paragraph });
  const native = storyRanges.rows.find(row => row.name === 'Headers-across-paragraphs')!;
  const snapshot = (instance: Editor) => instance.getJSON().content!.map(p => {
    const read = (text: string, attrs: Record<string, unknown> = {}) => ({ text,
      family: String(attrs.fontFamily || p.attrs!.paragraphFontFamily).replace(/^['"]|['"]$/g, ''),
      size: parseFloat(String(attrs.fontSize || p.attrs!.paragraphFontSize)),
    });
    const characters = p.content!.flatMap(node => {
      const text = node.type === 'wordTab' ? '\t' : 'text' in node ? node.text : '';
      expect(text).not.toBe('');
      return [...text].map(char => read(char, node.marks?.find(mark => mark.type === 'textStyle')?.attrs));
    });
    characters.push(read('\r'));
    return characters;
  });
  try {
    const before = editor.getJSON();
    editor.chain().setTextSelection({ from: 3, to: 7 })
      .setWordFontFamily('"Times New Roman"').setWordFontSize('20pt').run();
    expect(snapshot(editor)).toEqual(native.paragraphs.slice(0, 2).map(p => p.characters));
    const after = editor.getJSON();
    editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
    editor.commands.redo(); expect(editor.getJSON()).toEqual(after);
    const reloaded = new Editor({ extensions: wordExtensions(), content: editor.getHTML() });
    try { expect(snapshot(reloaded)).toEqual(snapshot(editor)); } finally { reloaded.destroy(); }
  } finally { editor.destroy(); }
});

const initial =
  '<p style="font-family:Arial;font-size:12pt">Alpha</p><p style="font-family:Arial;font-size:12pt">Beta</p>';

for (const key of ['size', 'family'] as const)
 for (const content of [initial, '<p style="font-family:Arial;font-size:12pt"><span style="font-family:Arial;font-size:12pt;color:red">Alpha</span><span data-word-tab="true" style="font-family:Arial;font-size:12pt">\t</span></p>'])
  it(`does not insert an empty undo step when the explicit ${key} is applied twice (${content === initial ? 'inherited' : 'marked'})`, () => {
    const editor = new Editor({ extensions: wordExtensions(), content });
    try {
      editor.commands.setTextSelection({ from: 1, to: editor.state.doc.firstChild!.content.size + 1 });
      const before = editor.getJSON();
      const apply = () => key === 'size'
        ? editor.commands.setWordFontSize('20pt')
        : editor.commands.setWordFontFamily('"Courier New"');
      apply();
      const after = editor.getJSON();
      apply();
      expect(editor.getJSON()).toEqual(after);
      editor.commands.undo();
      expect(editor.getJSON()).toEqual(before);
      editor.commands.redo();
      expect(editor.getJSON()).toEqual(after);
    } finally { editor.destroy(); }
  });

for (const [name, from, to, changedMark] of [
  ['whole paragraph text', 1, 6, true],
  ['paragraph mark included', 1, 8, true],
  ['end caret', 6, 6, true],
  ['first character', 1, 2, false],
  ['middle text', 2, 5, false],
  ['last text part', 3, 6, false],
  ['start caret', 1, 1, false],
  ['mark only', 6, 8, true],
] as const)
  it(`preserves native paragraph and unselected text fonts for ${name}`, () => {
    const editor = new Editor({ extensions: wordExtensions(), content: initial });
    try {
      editor.commands.setTextSelection({ from, to });
      const before = editor.getJSON();
      editor.chain().setWordFontFamily('Courier New').setWordFontSize('20pt').run();
      const after = editor.getJSON();
      expect(after.content![0].attrs).toMatchObject({
        paragraphFontFamily: changedMark ? 'Courier New' : 'Arial',
        paragraphFontSize: changedMark ? '20pt' : '12pt',
      });
      expect(after.content![1]).toEqual(before.content![1]);
      let offset = 1;
      for (const node of after.content![0].content!) {
        if (!('text' in node)) throw new Error('Expected a text run');
        for (const _ of node.text!) {
          const selected = offset >= from && offset < to;
          const style = node.marks?.find((m) => m.type === 'textStyle')?.attrs;
          expect(style?.fontSize || after.content![0].attrs!.paragraphFontSize).toBe(
            selected ? '20pt' : '12pt',
          );
          expect(style?.fontFamily || after.content![0].attrs!.paragraphFontFamily).toBe(
            selected ? 'Courier New' : 'Arial',
          );
          offset++;
        }
      }
      editor.commands.undo();
      expect(editor.getJSON()).toEqual(before);
      editor.commands.redo();
      expect(editor.getJSON()).toEqual(after);
      const reload = new Editor({ extensions: wordExtensions(), content: editor.getHTML() });
      expect(reload.getHTML()).toBe(editor.getHTML());
      reload.destroy();
    } finally {
      editor.destroy();
    }
  });

it('formats an ASCII word at an interior caret without moving it or changing the paragraph mark', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: initial });
  try {
    editor.commands.setTextSelection(3);
    const before = editor.getJSON();
    editor.chain().setWordFontFamily('Courier New').setWordFontSize('20pt').run();
    const after = editor.getJSON();
    expect(editor.state.selection.from).toBe(3);
    expect(editor.state.selection.empty).toBe(true);
    expect(after.content![0].attrs).toEqual(before.content![0].attrs);
    expect(
      after.content![0].content![0].marks?.find((m) => m.type === 'textStyle')?.attrs,
    ).toMatchObject({ fontFamily: 'Courier New', fontSize: '20pt' });
    expect(after.content![1]).toEqual(before.content![1]);
    editor.commands.undo();
    expect(editor.getJSON()).toEqual(before);
    editor.commands.redo();
    expect(editor.getJSON()).toEqual(after);
  } finally {
    editor.destroy();
  }
});

it('persists an empty paragraph font before typing and restores it through history', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p></p>' });
  try {
    editor.chain().setWordFontFamily('Arial').setWordFontSize('20pt').run();
    const empty = editor.getHTML();
    editor.commands.setContent(empty);
    editor.commands.insertContent('Typed');
    expect(
      editor.getJSON().content![0].content![0].marks?.find((m) => m.type === 'textStyle')?.attrs,
    ).toMatchObject({ fontFamily: 'Arial', fontSize: '20pt' });
    editor.commands.undo();
    expect(editor.getText()).toBe('');
    expect(editor.getJSON().content![0].attrs).toMatchObject({
      paragraphFontFamily: 'Arial',
      paragraphFontSize: '20pt',
    });
    editor.commands.redo();
    expect(editor.getText()).toBe('Typed');
  } finally {
    editor.destroy();
  }
});

it('rejects invalid font inputs and keeps can() checks free of model changes', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: initial });
  try {
    editor.commands.selectAll();
    const before = editor.getJSON();
    for (const size of ['0pt', '1639pt', '12.25pt', 'NaNpt', '12px'])
      expect(editor.commands.setWordFontSize(size)).toBe(false);
    expect(editor.commands.setWordFontFamily('')).toBe(false);
    expect(editor.can().setWordFontSize('20pt')).toBe(true);
    expect(editor.getJSON()).toEqual(before);
  } finally {
    editor.destroy();
  }
});
