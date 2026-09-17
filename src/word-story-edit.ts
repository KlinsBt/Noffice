import { Extension, type Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import type { WordContent } from './model';
import { wordStoriesSchema, type WordStories, type WordStoryReference } from './word-stories';
import { sanitizeHTML } from './formats';
import { wordJSON } from './word-extensions';
import { resolveWordSections } from './word-section-layout';

/** Story edits participate in the body editor's document-wide undo history.
 * This attribute is initialized from typed content, never parsed from HTML. */
export const WordStoryState = Extension.create({
  name: 'wordStoryState',
  addGlobalAttributes: () => [
    {
      types: ['doc'],
      attributes: {
        wordStories: { default: null, rendered: false, parseHTML: () => null },
      },
    },
  ],
});
export function initializeWordStories(editor: Editor, content: WordContent) {
  editor.view.dispatch(
    editor.state.tr
      .setDocAttribute('wordStories', content.stories || null)
      .setMeta('addToHistory', false)
      .setMeta('preventUpdate', true),
  );
}
export function changeWordStory(editor: Editor, path: string, expected: string, html: string) {
  const current = editor.state.doc.attrs.wordStories as WordStories | null;
  const part = current?.parts.find((part) => part.path === path);
  if (!current || !part || part.html !== expected)
    throw Error('This header or footer changed while it was open. Reopen it before editing.');
  if (html.length > 1000000) throw Error('This header or footer exceeds the supported size.');
  const clean = sanitizeHTML(html);
  if (JSON.stringify(wordJSON(part.html)) === JSON.stringify(wordJSON(clean))) return;
  const next = wordStoriesSchema.parse({
    ...current,
    parts: current.parts.map((p) => (p.path === path ? { ...p, html: clean } : p)),
  });
  editor.view.dispatch(closeHistory(editor.state.tr).setDocAttribute('wordStories', next));
  editor.view.dispatch(closeHistory(editor.state.tr));
}
export function wordStoryChoices(content: WordContent) {
  const choices: {
    path: string | null;
    label: string;
    active: boolean;
    sectionId: string;
    kind: WordStoryReference['kind'];
    slot: WordStoryReference['slot'];
  }[] = [];
  resolveWordSections(content).forEach((section, i) => {
    for (const [key, label] of [
      ['headers', 'header'],
      ['footers', 'footer'],
    ] as const)
      for (const slot of ['default', 'first', 'even'] as const) {
        const ref = section[key][slot];
        const part =
          ref && content.stories?.parts.find((p) => p.relationshipIds.includes(ref.relationshipId));
        const active =
          slot === 'default' ||
          (slot === 'first' ? section.differentFirstPage : !!content.stories?.evenAndOddHeaders);
        choices.push({
          path: part?.path || null,
          sectionId: section.id,
          kind: label,
          slot,
          active,
          label: `Section ${i + 1} — ${slot === 'first' ? 'First-page' : slot === 'even' ? 'Even-page' : 'Default'} ${label}${ref?.inherited ? ' (linked)' : ''}${part ? '' : ' (empty)'}${active ? '' : ' (inactive)'}`,
        });
      }
  });
  return choices;
}
