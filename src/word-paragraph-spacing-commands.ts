import { Extension } from '@tiptap/core';
import type { EditorState } from '@tiptap/pm/state';
import { closeHistory } from '@tiptap/pm/history';
import { selectedWordParagraphs } from './word-paragraph-selection';
import { wordParagraphSpace, wordParagraphSpaceAttributes, wordParagraphSpaceFromAttrs, type WordParagraphSpace } from './word-paragraph-spacing';

type Change = { before?: WordParagraphSpace; after?: WordParagraphSpace; contextual?: boolean };
declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    wordParagraphSpacing: { setWordParagraphSpacing: (change: Change) => ReturnType };
  }
}

export function wordSelectionParagraphSpacing(state: EditorState) {
  const paragraphs = selectedWordParagraphs(state);
  const common = <T>(read: (attrs: Record<string, unknown>) => T): T | null => {
    if (!paragraphs.length) return null;
    const first = read(paragraphs[0].attrs);
    return paragraphs.every(p => JSON.stringify(read(p.attrs)) === JSON.stringify(first)) ? first : null;
  };
  return { before: common(a => wordParagraphSpaceFromAttrs(a, 'before')),
    after: common(a => wordParagraphSpaceFromAttrs(a, 'after')),
    contextual: common(a => !!a.paragraphContextualSpacing) };
}

export const WordParagraphSpacingCommands = Extension.create({
  name: 'wordParagraphSpacingCommands',
  addCommands() {
    return { setWordParagraphSpacing: change => ({ tr, dispatch }) => {
      const attrs: Record<string, unknown> = {};
      for (const side of ['before', 'after'] as const) {
        if (change[side] === undefined) continue;
        const value = wordParagraphSpace(change[side]);
        if (!value) return false;
        Object.assign(attrs, wordParagraphSpaceAttributes(side, value));
      }
      if (change.contextual !== undefined) {
        if (typeof change.contextual !== 'boolean') return false;
        attrs.paragraphContextualSpacing = change.contextual;
      }
      const paragraphs = selectedWordParagraphs(tr);
      if (!paragraphs.length || !Object.keys(attrs).length) return false;
      const edits = paragraphs.filter(p => Object.entries(attrs).some(([key, value]) => JSON.stringify(p.attrs[key]) !== JSON.stringify(value)));
      if (!dispatch || !edits.length) return true;
      closeHistory(tr);
      for (const p of edits) tr.setNodeMarkup(p.pos, undefined, { ...p.attrs, ...attrs });
      return true;
    } };
  },
});
