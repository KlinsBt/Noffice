import { expect, it } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { WordStoryState, initializeWordStories, changeWordStory } from './word-story-edit';
import { newFile } from './model';

it('joins story changes and ordinary typing to one ordered document history', () => {
  const file = newFile('word');
  if (file.content.kind !== 'word') throw Error('Expected Word');
  file.content.stories = {
    version: 1,
    evenAndOddHeaders: false,
    parts: [
      {
        path: 'word/header1.xml',
        kind: 'header',
        relationshipIds: ['header'],
        html: '<p>Original</p>',
      },
    ],
  };
  const editor = new Editor({
    extensions: [...wordExtensions(), WordStoryState],
    content: '<p>Body</p>',
  });
  try {
    initializeWordStories(editor, file.content);
    expect(editor.can().undo()).toBe(false);
    changeWordStory(editor, 'word/header1.xml', '<p>Original</p>', '<p>Edited</p>');
    const story = () => editor.state.doc.attrs.wordStories.parts[0].html;
    expect(story()).toBe('<p>Edited</p>');
    editor.commands.setTextSelection(5);
    editor.commands.insertContent(' typed');
    editor.commands.undo();
    expect(editor.getText()).toBe('Body');
    expect(story()).toBe('<p>Edited</p>');
    editor.commands.undo();
    expect(story()).toBe('<p>Original</p>');
    editor.commands.redo();
    editor.commands.redo();
    expect(editor.getText()).toBe('Body typed');
    expect(story()).toBe('<p>Edited</p>');
    expect(editor.getHTML()).not.toContain('header1');
    const before = editor.state.doc;
    expect(() =>
      changeWordStory(editor, 'word/header1.xml', '<p>Original</p>', '<p>Stale</p>'),
    ).toThrow('changed while');
    expect(editor.state.doc).toBe(before);
    changeWordStory(editor, 'word/header1.xml', '<p>Edited</p>', '<p>Edited</p>');
    expect(editor.state.doc).toBe(before);
  } finally {
    editor.destroy();
  }
});
