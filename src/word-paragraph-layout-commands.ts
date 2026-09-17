import { Extension } from '@tiptap/core';
import type { EditorState } from '@tiptap/pm/state';
import { closeHistory } from '@tiptap/pm/history';
import { paragraphIndents, paragraphBreaks } from './paragraph-layout';
import { selectedWordParagraphs } from './word-paragraph-selection';

export type WordParagraphAlignment = 'left' | 'center' | 'right' | 'justify';
export type WordParagraphLayout = {
  textAlign: WordParagraphAlignment;
  indentStart: number; indentEnd: number; firstLineIndent: number;
  keepNext: boolean; keepLines: boolean; pageBreakBefore: boolean; widowControl: boolean;
};
const indents = new Set<string>(paragraphIndents.map(([name]) => name));
const breaks = new Set<string>(paragraphBreaks.map(([name]) => name));
const alignments = ['left', 'center', 'right', 'justify'];

/** Independently measured native physical-indent entry: Single precision,
 * twips rounded away from zero at ties, then the signed 1,584pt limit.
 * Character-unit authoring remains a separate requirement. */
export function wordParagraphIndentInput(points: unknown): number | null {
  if (typeof points !== 'number' || !Number.isFinite(points)) return null;
  const native = Math.fround(points);
  const twips = Math.sign(native) * Math.round(Math.abs(native) * 20);
  return Number.isSafeInteger(twips) && Math.abs(twips) <= 31680 ? (twips === 0 ? 0 : twips / 20) : null;
}
function read(attrs: Record<string, unknown>, key: keyof WordParagraphLayout): unknown {
  if (indents.has(key)) {
    const value = attrs[key];
    if (value == null || value === '') return 0;
    const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))(pt|px)?$/.exec(String(value));
    return match ? Number(match[1]) * (match[2] === 'px' ? .75 : 1) : null;
  }
  if (key === 'textAlign') return attrs[key] ?? 'left';
  return attrs[key] ?? (key === 'widowControl');
}
export function wordSelectionParagraphLayout(state: EditorState): {
  [K in keyof WordParagraphLayout]: WordParagraphLayout[K] | null
} {
  const paragraphs = selectedWordParagraphs(state);
  return Object.fromEntries(['textAlign', ...indents, ...breaks].map(name => {
    const key = name as keyof WordParagraphLayout;
    const first = paragraphs.length ? read(paragraphs[0].attrs, key) : null;
    return [key, paragraphs.every(p => read(p.attrs, key) === first) ? first : null];
  })) as ReturnType<typeof wordSelectionParagraphLayout>;
}
declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    wordParagraphLayout: { setWordParagraphLayout: (patch: Partial<WordParagraphLayout>) => ReturnType };
  }
}
export const WordParagraphLayoutCommands = Extension.create({
  name: 'wordParagraphLayoutCommands',
  priority: 110,
  addKeyboardShortcuts() {
    return {
      'Mod-Shift-l': () => this.editor.commands.setWordParagraphLayout({ textAlign: 'left' }),
      'Mod-Shift-e': () => this.editor.commands.setWordParagraphLayout({ textAlign: 'center' }),
      'Mod-Shift-r': () => this.editor.commands.setWordParagraphLayout({ textAlign: 'right' }),
      'Mod-Shift-j': () => this.editor.commands.setWordParagraphLayout({ textAlign: 'justify' }),
    };
  },
  addCommands() {
    return { setWordParagraphLayout: patch => ({ tr, dispatch }) => {
      if (!patch || typeof patch !== 'object' || Array.isArray(patch)) return false;
      const entries = Object.entries(patch);
      if (!entries.length) return false;
      const values: [keyof WordParagraphLayout, unknown, unknown][] = [];
      for (const [key, value] of entries) {
        let semantic: unknown = value, stored: unknown = value;
        if (indents.has(key)) {
          const points = wordParagraphIndentInput(value);
          if (points === null) return false;
          semantic = points; stored = `${points}pt`;
        } else if (breaks.has(key)) {
          if (typeof value !== 'boolean') return false;
        } else if (key !== 'textAlign' || typeof value !== 'string' || !alignments.includes(value)) return false;
        values.push([key as keyof WordParagraphLayout, semantic, stored]);
      }
      const paragraphs = selectedWordParagraphs(tr);
      if (!paragraphs.length) return false;
      const edits = paragraphs.map(p => ({ ...p,
        changes: values.filter(([key, value]) => read(p.attrs, key) !== value),
      })).filter(p => p.changes.length);
      if (!dispatch || !edits.length) return true;
      closeHistory(tr);
      for (const p of edits) tr.setNodeMarkup(p.pos, undefined, { ...p.attrs,
        ...Object.fromEntries(p.changes.map(([key, , value]) => [key, value])),
      });
      return true;
    } };
  },
});
