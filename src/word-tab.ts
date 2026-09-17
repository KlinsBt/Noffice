import { Node } from '@tiptap/core';

/** An inline tab must survive HTML whitespace normalization as a distinct Word character. */
export const WordTab = Node.create({
  name: 'wordTab',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: false,
  parseHTML: () => [{ tag: 'span[data-word-tab]' }],
  renderHTML: () => [
    'span',
    { 'data-word-tab': 'true', style: 'white-space:pre;tab-size:4' },
    '\t',
  ],
  renderText: () => '\t',
  addKeyboardShortcuts() {
    return {
      Tab: () =>
        this.editor.isActive('table')
          ? false
          : this.editor.commands.insertContent({ type: 'wordTab' }),
    };
  },
});
