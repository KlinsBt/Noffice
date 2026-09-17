import { Extension, type Editor } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import type { DocxStructure } from './docx-sections';
import type { WordStories } from './word-stories';
import { repairWordStorySections } from './word-story-section-repair';
import { mapWordSectionRanges, type WordSectionMap } from './word-section-ranges';
import {
  initializeSectionState,
  mapLiveSections,
  transformSectionState,
  validateSectionState,
  type WordSectionState,
} from './word-section-breaks';

const equal = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
function currentMap(doc: Editor['state']['doc'], source: DocxStructure | undefined) {
  const live = doc.attrs.wordSectionState as WordSectionState | null;
  return live && (source || live.version === 2)
    ? mapLiveSections(doc, source, live)
    : mapWordSectionRanges(doc, source);
}

const key = new PluginKey<{ source: DocxStructure | undefined; map: WordSectionMap }>(
  'wordSections',
);
/** Cache document positions on document transactions, not on every cursor movement. */
export const WordEditorSections = Extension.create<{ source: DocxStructure | undefined }>({
  name: 'wordEditorSections',
  addOptions: () => ({ source: undefined }),
  addGlobalAttributes() {
    return [
      {
        types: ['doc'],
        attributes: { wordSectionState: { default: null, rendered: false, parseHTML: () => null } },
      },
    ];
  },
  addProseMirrorPlugins() {
    const source = this.options.source;
    return [
      new Plugin({
        key,
        state: {
          init: (_, state) => ({ source, map: currentMap(state.doc, source) }),
          apply: (tr, previous) => {
            const update = tr.getMeta(key) as { source: DocxStructure | undefined } | undefined;
            if (!tr.docChanged && !update) return previous;
            const source = update ? update.source : previous.source;
            // Document edits are followed by the boundary-repair transaction below.
            // Its new ordinals must be installed before consumers resolve the final map.
            try {
              return { source, map: currentMap(tr.doc, source) };
            } catch {
              return { source, map: mapWordSectionRanges(tr.doc, source) };
            }
          },
        },
        appendTransaction(transactions, oldState, newState) {
          const live = oldState.doc.attrs.wordSectionState as WordSectionState | null;
          if (
            !live ||
            !transactions.some((tr) => tr.docChanged) ||
            !equal(live, newState.doc.attrs.wordSectionState)
          )
            return null;
          const next = transformSectionState(oldState.doc, newState.doc, live, transactions);
          if (equal(live, next)) return null;
          const tr = newState.tr.setDocAttribute('wordSectionState', next);
          const source = key.getState(newState)?.source;
          const stories = newState.doc.attrs.wordStories as WordStories | null;
          if (
            (source || live.version === 2) &&
            stories &&
            equal(stories, oldState.doc.attrs.wordStories)
          ) {
            const repaired = repairWordStorySections(source, stories, live, next);
            if (!equal(repaired, stories)) tr.setDocAttribute('wordStories', repaired);
          }
          return tr;
        },
      }),
    ];
  },
});

export const wordSectionMap = (state: Editor['state']) => key.getState(state)?.map;

export function updateWordSectionSource(
  editor: Editor,
  source: DocxStructure | undefined,
  saved?: WordSectionState,
) {
  if (!editor.state.doc.attrs.wordSectionState && (source || saved?.version === 2)) {
    const live = saved || initializeSectionState(editor.state.doc, source);
    if (live) validateSectionState(live, source, editor.state.doc);
    if (live)
      editor.view.dispatch(
        editor.state.tr
          .setDocAttribute('wordSectionState', live)
          .setMeta('addToHistory', false)
          .setMeta('preventUpdate', true),
      );
  }
  if (key.getState(editor.state)?.source === source) return;
  editor.view.dispatch(editor.state.tr.setMeta(key, { source }).setMeta('addToHistory', false));
}
