import { it, expect } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { WordPageLayout, initializeWordPageLayout, changeWordPageLayout } from './word-page-layout';
import { newFile, contentSchema } from './model';

it('joins global page commands to undo/redo without mixing the following text edit', () => {
  const content = newFile('word').content;
  if (content.kind !== 'word') throw Error();
  const editor = new Editor({
    extensions: [...wordExtensions(), WordPageLayout],
    content: '<p>Hello</p>',
  });
  try {
    initializeWordPageLayout(editor, content);
    expect(editor.can().undo()).toBe(false);
    changeWordPageLayout(editor, 'paper', 'letter');
    editor.commands.insertContent('Text');
    editor.commands.undo();
    expect(editor.getText()).toBe('Hello');
    expect(editor.state.doc.attrs.wordPageLayout.paper).toBe('letter');
    editor.commands.undo();
    expect(editor.state.doc.attrs.wordPageLayout).toEqual({
      paper: 'a4',
      margin: 'normal',
      orientation: undefined,
      pageOverrides: undefined,
    });
    editor.commands.redo();
    expect(editor.state.doc.attrs.wordPageLayout.pageOverrides).toEqual({ paper: true });
    expect(editor.getHTML()).not.toContain('pageOverrides');
    expect(contentSchema.parse({ ...content, ...editor.state.doc.attrs.wordPageLayout }).kind).toBe(
      'word',
    );
  } finally {
    editor.destroy();
  }
});
it('retains explicit same-value overrides through undo and ignores repeated identical commands', () => {
  const content = newFile('word').content;
  if (content.kind !== 'word') throw Error();
  const editor = new Editor({ extensions: [...wordExtensions(), WordPageLayout], content: '<p/>' });
  try {
    initializeWordPageLayout(editor, content);
    changeWordPageLayout(editor, 'paper', 'a4');
    changeWordPageLayout(editor, 'paper', 'a4');
    expect(editor.state.doc.attrs.wordPageLayout.pageOverrides).toEqual({ paper: true });
    editor.commands.undo();
    expect(editor.state.doc.attrs.wordPageLayout.pageOverrides).toBeUndefined();
    expect(editor.can().undo()).toBe(false);
  } finally {
    editor.destroy();
  }
});
