import { Editor } from '@tiptap/core';
import { afterEach, expect, it } from 'vitest';
import { wordExtensions } from './word-extensions';
import { wordMatches, replaceWordMatches } from './word-search';
let editor: Editor;
afterEach(() => editor?.destroy());

it('finds across formatting boundaries and replaces with source identities and undo intact', () => {
  editor = new Editor({
    extensions: wordExtensions(),
    content:
      '<p data-source-paragraph="0:0"><strong>Project </strong><em>Atlas</em> launch</p><p>Project Atlas</p>',
  });
  const matches = wordMatches(editor, 'Project Atlas');
  expect(matches).toHaveLength(2);
  replaceWordMatches(editor, matches, 'Plan');
  expect(editor.getText()).toBe('Plan launch\n\nPlan');
  expect(editor.getJSON().content![0].attrs!.sourceParagraph).toBe('0:0');
  expect(editor.getHTML()).toContain('<strong>Plan</strong>');
  editor.commands.undo();
  expect(wordMatches(editor, 'Project Atlas')).toHaveLength(2);
  editor.commands.redo();
  expect(editor.getText()).toBe('Plan launch\n\nPlan');
});
it('does not invent matches through paragraph boundaries or inline objects', () => {
  editor = new Editor({
    extensions: wordExtensions(),
    content: '<p>one</p><p>two</p><p>one<br>two</p>',
  });
  expect(wordMatches(editor, 'onetwo')).toEqual([]);
  replaceWordMatches(editor, wordMatches(editor, 'one'), '');
  expect(wordMatches(editor, 'one')).toEqual([]);
  expect(wordMatches(editor, 'two')).toHaveLength(2);
});
