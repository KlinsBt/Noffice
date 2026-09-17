import { Extension } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    wordBoundaryDelete: {
      /** Delete an exactly selected top-level paragraph/section boundary. */
      deleteWordBoundary: (smartSpacing?: boolean) => ReturnType;
    };
  }
}

/** Word's selected-boundary command differs from a raw model deletion. Its
 * measured ASCII punctuation spacing belongs to the user command, not to
 * arbitrary transforms, import, undo or structural repair transactions. */
export const WordBoundaryDelete = Extension.create({
  name: 'wordBoundaryDelete',
  priority: 1100,
  addCommands() {
    return {
      deleteWordBoundary: (smartSpacing = true) => ({ state, tr, dispatch }) => {
        const { selection } = state;
        if (!(selection instanceof TextSelection) || selection.empty) return false;
        const { $from: a, $to: b, from, to } = selection;
        if (a.depth !== 1 || b.depth !== 1 || a.parent.type.name !== 'paragraph' ||
            b.parent.type.name !== 'paragraph' || a.parentOffset !== a.parent.content.size ||
            b.parentOffset !== 0 || b.index(0) !== a.index(0) + 1) return false;
        const left = a.parent.lastChild, right = b.parent.firstChild;
        // Only the measured plain ASCII letter/punctuation contexts are
        // qualified here. Inline controls and other scripts retain raw joins.
        const space = smartSpacing && left?.isText && right?.isText && (
          (/[!),.:;?\]]$/.test(left.text || '') && /^[A-Za-z]/.test(right.text || '')) ||
          (/[A-Za-z]$/.test(left.text || '') && /^[([]/.test(right.text || ''))
        );
        if (dispatch) {
          // Native insertion takes the following run's formatting, including
          // a font distinct from either the preceding run or paragraph mark.
          if (space) tr.replaceWith(from, to, state.schema.text(' ', right!.marks));
          else tr.delete(from, to);
          // The first paragraph keeps its layout. Its deleted mark takes the
          // following paragraph mark's font and retained run properties.
          const attrs: Record<string, unknown> = { ...a.parent.attrs,
            paragraphMarkSource: b.parent.attrs.paragraphMarkSource || b.parent.attrs.sourceParagraph || 'new',
          };
          for (const key of ['paragraphFontSize', 'paragraphFontFamily', 'paragraphKerning',
            'paragraphFontFeatures', 'paragraphScript']) attrs[key] = b.parent.attrs[key];
          tr.setNodeMarkup(a.before(), undefined, attrs);
          // Native caret remains before the automatically supplied space.
          tr.setSelection(TextSelection.create(tr.doc, from)).scrollIntoView();
        }
        return true;
      },
    };
  },
  addKeyboardShortcuts() {
    return {
      Backspace: () => this.editor.commands.deleteWordBoundary(),
      Delete: () => this.editor.commands.deleteWordBoundary(),
    };
  },
});
