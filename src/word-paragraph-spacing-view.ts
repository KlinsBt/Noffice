import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import type { Node } from '@tiptap/pm/model';

const paragraph = (node?: Node | null) => !!node && ['paragraph', 'heading'].includes(node.type.name);
const style = (node: Node) => node.attrs.paragraphStyleId || (node.type.name === 'heading' ? `Heading${node.attrs.level}` : 'Normal');

/** Same-style suppression belongs to the rendered adjacency, not saved margins. */
export const WordParagraphSpacingView = Extension.create({
  name: 'wordParagraphSpacingView',
  addProseMirrorPlugins() {
    return [new Plugin({ props: { decorations(state) {
      const decorations: Decoration[] = [];
      state.doc.descendants((node, pos, parent, index) => {
        if (!paragraph(node) || !node.attrs.paragraphContextualSpacing || !parent) return;
        const previous = index > 0 ? parent.child(index - 1) : null;
        const next = index + 1 < parent.childCount ? parent.child(index + 1) : null;
        const before = paragraph(previous) && style(previous!) === style(node);
        const after = paragraph(next) && style(next!) === style(node);
        if (before || after) decorations.push(Decoration.node(pos, pos + node.nodeSize, {
          class: [before ? 'word-contextual-before' : '', after ? 'word-contextual-after' : ''].filter(Boolean).join(' '),
        }));
      });
      return DecorationSet.create(state.doc, decorations);
    } } })];
  },
});
