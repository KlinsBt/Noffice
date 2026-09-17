import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { isFlatWordList } from './word-flat-list';
import { sectionParagraphs, type WordSectionState } from './word-section-breaks';
import { resolveWordLists } from './word-list-resolve';
import type { WordNumbering } from './word-list-layout';

/** Word's empty section-ending list paragraph neither paints a marker nor
 * increments numbering. Values are derived view decorations, never saved HTML:
 * typing, deleting, history and reload must recalculate the same live sequence. */
export const WordListSectionView = Extension.create({
  name: 'wordListSectionView',
  addOptions() { return { numbering: (): WordNumbering | undefined => undefined }; },
  addProseMirrorPlugins() {
    const options = this.options;
    return [new Plugin({
      props: {
        decorations(state) {
          const resolved = resolveWordLists(state.doc, options.numbering());
          if (resolved?.size) {
            const decorations: Decoration[] = [], lists = new Set<number>();
            for (const p of resolved.values()) {
              if (!lists.has(p.listFrom)) {
                decorations.push(Decoration.node(p.listFrom, p.listTo, { class: 'word-source-list' }));
                lists.add(p.listFrom);
              }
              decorations.push(Decoration.node(p.from, p.to, {
                'data-word-list-marker': p.marker,
                style: `margin-inline-start:${p.definition.left / 20}pt;text-indent:0;` +
                  `--word-list-marker-left:${-p.definition.hanging / 20}pt;--word-list-marker-size:${p.definition.size}pt;` +
                  `--word-list-marker-font:"${p.definition.font}"`,
              }));
              decorations.push(Decoration.node(p.from - 1, p.to + 1, p.marker
                ? { ...(p.definition.format === 'decimal' ? { value: String(p.value) } : {}) }
                : { style: 'list-style-type:none', 'data-word-section-list-empty': 'true' }));
            }
            return DecorationSet.create(state.doc, decorations);
          }
          const sections = state.doc.attrs.wordSectionState as WordSectionState | null;
          if (!sections?.breaks.length) return null;
          const paragraphs = sectionParagraphs(state.doc);
          const hidden = new Set(sections.breaks.map(b => paragraphs[b.paragraph])
            .filter(p => p && !p.node.content.size).map(p => p.from));
          if (!hidden.size) return null;
          const decorations: Decoration[] = [];
          state.doc.forEach((list, from) => {
            if (!isFlatWordList(list)) return;
            let value = list.type.name === 'orderedList' ? list.attrs.start : 1;
            list.forEach((item, offset) => {
              const position = from + 1 + offset;
              const suppress = hidden.has(position + 1);
              decorations.push(Decoration.node(position, position + item.nodeSize, suppress
                ? { style: 'list-style-type:none', 'data-word-section-list-empty': 'true' }
                : list.type.name === 'orderedList' ? { value: String(value) } : {}));
              if (!suppress) value++;
            });
          });
          return DecorationSet.create(state.doc, decorations);
        },
      },
    })];
  },
});
