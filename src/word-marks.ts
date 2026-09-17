import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { wordSectionMap } from './word-editor-sections';
const key = new PluginKey<boolean>('wordFormattingMarks');
export const WordFormattingMarks = Extension.create({
  name: 'wordFormattingMarks',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key,
        state: { init: () => false, apply: (tr, previous) => tr.getMeta(key) ?? previous },
        props: {
          decorations(state) {
            if (!key.getState(state)) return DecorationSet.empty;
            const decorations: Decoration[] = [];
            const sections = wordSectionMap(state);
            const sectionEnds = new Set(sections?.ranges.slice(0, -1).map((r) => r.to));
            const marker = (pos: number, symbol: string, kind: string) => {
              decorations.push(
                Decoration.widget(
                  pos,
                  () => {
                    const span = document.createElement('span');
                    span.className = `word-format-mark ${kind}`;
                    span.dataset.symbol = symbol;
                    span.setAttribute('aria-hidden', 'true');
                    span.contentEditable = 'false';
                    return span;
                  },
                  { side: -1 },
                ),
              );
            };
            state.doc.descendants((node, pos) => {
              if (node.isText)
                for (let i = 0; i < node.text!.length; i++)
                  if (node.text![i] === ' ')
                    decorations.push(
                      Decoration.inline(pos + i, pos + i + 1, { class: 'word-space-mark' }),
                    );
              if (['paragraph', 'heading'].includes(node.type.name)) {
                const isSectionEnd = sectionEnds.has(pos + node.nodeSize);
                marker(
                  pos + node.nodeSize - 1,
                  isSectionEnd ? 'Section break' : '¶',
                  isSectionEnd ? 'section-mark' : 'paragraph-mark',
                );
              }
              if (node.type.name === 'wordTab') marker(pos, '→', 'tab-mark');
              if (node.type.name === 'hardBreak') marker(pos, '↵', 'break-mark');
            });
            return DecorationSet.create(state.doc, decorations);
          },
        },
      }),
    ];
  },
});
export function showFormattingMarks(editor: import('@tiptap/core').Editor, show: boolean) {
  editor.view.dispatch(editor.state.tr.setMeta(key, show).setMeta('addToHistory', false));
}
