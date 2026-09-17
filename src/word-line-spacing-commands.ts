import { Extension } from '@tiptap/core';
import type { EditorState } from '@tiptap/pm/state';
import { closeHistory } from '@tiptap/pm/history';
import { selectedWordParagraphs } from './word-paragraph-selection';
import { wordLineSpacing, type WordLineSpacing } from './word-line-spacing';
import { wordDefaults } from './word-defaults';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    wordLineSpacing: { setWordLineSpacing: (spacing: WordLineSpacing) => ReturnType };
  }
}
const read = (attrs: Record<string, unknown>) => wordLineSpacing(
  attrs.paragraphLineHeight || String(wordDefaults.lineMultiple), attrs.paragraphLineRule);
const equal = (a: WordLineSpacing | null, b: WordLineSpacing | null) => a?.rule === b?.rule && a?.line === b?.line;

export function wordSelectionLineSpacing(state: EditorState): WordLineSpacing | null {
  const paragraphs = selectedWordParagraphs(state);
  if (!paragraphs.length) return null;
  const first = read(paragraphs[0].attrs);
  return paragraphs.every(p => equal(read(p.attrs), first)) ? first : null;
}

export const WordLineSpacingCommands = Extension.create({
  name: 'wordLineSpacingCommands',
  addCommands() {
    return { setWordLineSpacing: spacing => ({ tr, dispatch }) => {
      if (!spacing || !['auto', 'exact', 'atLeast'].includes(spacing.rule) ||
          !Number.isSafeInteger(spacing.line) || spacing.line < 14 || spacing.line > 31680) return false;
      const paragraphs = selectedWordParagraphs(tr);
      if (!paragraphs.length) return false;
      const edits = paragraphs.filter(p => !equal(read(p.attrs), spacing));
      if (!dispatch || !edits.length) return true;
      closeHistory(tr);
      const attrs = { paragraphLineHeight: spacing.rule === 'auto' ? String(spacing.line / 240) : `${spacing.line / 20}pt`,
        paragraphLineRule: spacing.rule === 'atLeast' ? 'atLeast' : null };
      for (const p of edits) tr.setNodeMarkup(p.pos, undefined, { ...p.attrs, ...attrs });
      return true;
    } };
  },
});
