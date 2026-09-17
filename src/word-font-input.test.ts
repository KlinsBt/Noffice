import { expect, it } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions, wordJSON } from './word-extensions';

for (const threshold of [0, 2, 24])
  it(`types the empty paragraph kerning threshold ${threshold} with atomic history and reload`, () => {
    const editor = new Editor({
      extensions: wordExtensions(),
      content: `<p data-word-paragraph-kerning="${threshold}" style="font-size:11pt;font-family:Calibri"></p>`,
    });
    const attrs = () =>
      editor
        .getJSON()
        .content![0].content?.at(-1)
        ?.marks?.find((m) => m.type === 'textStyle')?.attrs;
    try {
      editor.commands.insertContent('AV To');
      expect(attrs()?.wordKerning).toBe(threshold);
      editor.commands.undo();
      expect(editor.getText()).toBe('');
      editor.commands.redo();
      expect(attrs()?.wordKerning).toBe(threshold);
      const restored = new Editor({ extensions: wordExtensions(), content: editor.getHTML() });
      try {
        expect(
          restored.getJSON().content![0].content![0].marks?.find((m) => m.type === 'textStyle')
            ?.attrs?.wordKerning,
        ).toBe(threshold);
        expect(wordJSON(restored.getHTML())).toEqual(wordJSON(editor.getHTML()));
      } finally {
        restored.destroy();
      }
    } finally {
      editor.destroy();
    }
  });

it('takes terminal hard-line kerning from the paragraph mark and retains explicit typing overrides', () => {
  const initial =
    '<p data-word-paragraph-kerning="2" style="font-size:11pt;font-family:Calibri"><span data-word-kerning="24">Before<br></span></p>';
  const editor = new Editor({ extensions: wordExtensions(), content: initial });
  const last = () =>
    editor
      .getJSON()
      .content![0].content!.at(-1)!
      .marks?.find((m) => m.type === 'textStyle')?.attrs;
  try {
    editor.commands.setTextSelection(editor.state.doc.content.size - 1);
    editor.commands.insertContent('After');
    expect(last()?.wordKerning).toBe(2);
    expect(
      editor.getJSON().content![0].content![0].marks?.find((m) => m.type === 'textStyle')?.attrs
        ?.wordKerning,
    ).toBe(24);
    editor.commands.undo();
    expect(editor.getText()).not.toContain('After');
    editor.commands.redo();
    expect(last()?.wordKerning).toBe(2);
    editor.commands.setContent(initial);
    editor.commands.setTextSelection(editor.state.doc.content.size - 1);
    editor.commands.setMark('textStyle', { wordKerning: 0 });
    editor.commands.insertContent('Explicit');
    expect(last()?.wordKerning).toBe(0);
  } finally {
    editor.destroy();
  }
});

for (const kind of ['page', 'column'] as const)
  it(`retains the paragraph font on a new ${kind} break and terminal typing`, () => {
    const editor = new Editor({
      extensions: wordExtensions(),
      content: '<p style="font-size:10pt;font-family:Arial">before</p>',
    });
    try {
      editor.commands.setTextSelection(editor.state.doc.content.size - 1);
      editor.commands.insertWordFlowBreak(kind);
      const last = () => editor.getJSON().content!.at(-1)!.content!.at(-1)!;
      const flow = () =>
        editor
          .getJSON()
          .content!.flatMap((p) => p.content ?? [])
          .filter((n) => n.type === (kind === 'page' ? 'wordPageBreak' : 'wordColumnBreak'))
          .at(-1)!;
      expect(flow().marks?.find((m) => m.type === 'textStyle')?.attrs).toMatchObject({
        fontSize: '10pt',
        fontFamily: 'Arial',
      });
      editor.commands.insertContent('!');
      expect(last().marks?.find((m) => m.type === 'textStyle')?.attrs).toMatchObject({
        fontSize: '10pt',
        fontFamily: 'Arial',
      });
      editor.commands.undo();
      editor.commands.redo();
      expect(last().marks?.find((m) => m.type === 'textStyle')?.attrs).toMatchObject({
        fontSize: '10pt',
        fontFamily: 'Arial',
      });
      editor.commands.setFontSize('18pt');
      editor.commands.insertWordFlowBreak(kind);
      expect(flow().marks?.find((m) => m.type === 'textStyle')?.attrs?.fontSize).toBe('18pt');
    } finally {
      editor.destroy();
    }
  });

it('types on an empty final hard-break line using its paragraph mark, without changing preceding text', () => {
  const editor = new Editor({
    extensions: wordExtensions(),
    content:
      '<p style="font-size:40pt;font-family:Arial"><span style="font-size:10pt;font-family:Arial">Alpha<br></span></p>',
  });
  try {
    editor.commands.setTextSelection(editor.state.doc.content.size - 1);
    editor.commands.insertContent('New');
    const content = () => editor.getJSON().content![0].content!;
    expect(content()[0].marks?.find((m) => m.type === 'textStyle')?.attrs?.fontSize).toBe('10pt');
    expect(
      content()
        .at(-1)!
        .marks?.find((m) => m.type === 'textStyle')?.attrs?.fontSize,
    ).toBe('40pt');
    editor.commands.undo();
    expect(editor.getText()).not.toContain('New');
    editor.commands.redo();
    expect(
      content()
        .at(-1)!
        .marks?.find((m) => m.type === 'textStyle')?.attrs?.fontSize,
    ).toBe('40pt');
  } finally {
    editor.destroy();
  }
});

it('preserves explicit typing and rich paste fonts on an empty terminal line', () => {
  const initial =
    '<p style="font-size:40pt;font-family:Arial"><span style="font-size:10pt;font-family:Arial">Alpha<br></span></p>';
  const editor = new Editor({ extensions: wordExtensions(), content: initial });
  const lastFont = () =>
    editor
      .getJSON()
      .content![0].content!.at(-1)!
      .marks?.find((m) => m.type === 'textStyle')?.attrs;
  try {
    editor.commands.setTextSelection(editor.state.doc.content.size - 1);
    editor.commands.setFontSize('18pt');
    editor.commands.insertContent('Typed');
    expect(lastFont()?.fontSize).toBe('18pt');
    editor.commands.setContent(initial);
    editor.commands.setTextSelection(editor.state.doc.content.size - 1);
    const mark = editor.schema.marks.textStyle.create({ fontSize: '10pt', fontFamily: 'Arial' });
    editor.view.dispatch(
      editor.state.tr
        .replaceSelectionWith(editor.schema.text('Pasted', [mark]), false)
        .setMeta('uiEvent', 'paste'),
    );
    expect(lastFont()).toMatchObject({ fontSize: '10pt', fontFamily: 'Arial' });
    editor.commands.undo();
    expect(editor.getText()).not.toContain('Pasted');
    editor.commands.redo();
    expect(lastFont()?.fontSize).toBe('10pt');
  } finally {
    editor.destroy();
  }
});

it('stores the empty paragraph font on typed text in the same undo event', () => {
  const editor = new Editor({
    extensions: wordExtensions(),
    content: '<p style="font-size:30pt;font-family:Arial"></p>',
  });
  try {
    editor.commands.insertContent('Typed');
    const text = () => editor.getJSON().content![0].content![0];
    expect(text().marks?.find((m) => m.type === 'textStyle')?.attrs).toMatchObject({
      fontSize: '30pt',
      fontFamily: 'Arial',
    });
    editor.commands.undo();
    expect(editor.getText()).toBe('');
    editor.commands.redo();
    expect(text().marks?.find((m) => m.type === 'textStyle')?.attrs?.fontSize).toBe('30pt');
  } finally {
    editor.destroy();
  }
});

it('keeps an explicit typing size and rich pasted font over the paragraph fallback', () => {
  const editor = new Editor({
    extensions: wordExtensions(),
    content: '<p style="font-size:30pt;font-family:Arial"></p>',
  });
  try {
    editor.commands.setFontSize('18pt');
    editor.commands.insertContent('Typed');
    expect(
      editor.getJSON().content![0].content![0].marks?.find((m) => m.type === 'textStyle')?.attrs
        ?.fontSize,
    ).toBe('18pt');
    editor.commands.setContent('<p style="font-size:30pt;font-family:Arial"></p>');
    editor.commands.insertContent('<span style="font-size:14pt;font-family:Georgia">Pasted</span>');
    expect(
      editor.getJSON().content![0].content![0].marks?.find((m) => m.type === 'textStyle')?.attrs,
    ).toMatchObject({ fontSize: '14pt', fontFamily: 'Georgia' });
  } finally {
    editor.destroy();
  }
});
