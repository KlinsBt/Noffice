import { Node } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { closeHistory } from '@tiptap/pm/history';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    wordFlowBreak: { insertWordFlowBreak: (kind: 'page' | 'column') => ReturnType };
  }
}

/** A page boundary stays inside its source paragraph instead of inventing new paragraphs. */
export const WordPageBreak = Node.create({
  name: 'wordPageBreak',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  parseHTML: () => [{ tag: 'span[data-word-page-break]' }],
  renderHTML: () => [
    'span',
    {
      'data-word-page-break': 'true',
      'aria-label': 'Page break',
      contenteditable: 'false',
      style: 'display:block;break-after:page',
    },
    ['br'],
  ],
  renderText: () => '\f',
  addCommands() {
    return {
      insertWordFlowBreak:
        (kind) =>
        ({ state, tr, dispatch }) => {
          if (
            !state.selection.$from.sameParent(state.selection.$to) ||
            state.selection.$from.parent.type.name !== 'paragraph'
          )
            return false;
          const marks = [...(state.storedMarks ?? state.selection.$from.marks())];
          const type = state.schema.marks.textStyle;
          const inherited = type?.isInSet(marks);
          const paragraph = state.selection.$from.parent.attrs;
          if (type) {
            const attrs = { ...inherited?.attrs };
            attrs.fontSize ||= paragraph.paragraphFontSize;
            attrs.fontFamily ||= paragraph.paragraphFontFamily;
            if (inherited) marks.splice(marks.indexOf(inherited), 1);
            if (attrs.fontSize || attrs.fontFamily) marks.push(type.create(attrs));
          }
          if (!dispatch) return true;
          closeHistory(tr);
          tr.deleteSelection();
          let at = tr.selection.from;
          const $at = tr.doc.resolve(at);
          // Native keyboard insertion starts a new paragraph when there is
          // preceding content. Keep the original identity on the left; section
          // boundaries follow the mapped paragraph end onto the right.
          if ($at.parentOffset) {
            tr.split(at, 1, [
              {
                type: $at.parent.type,
                attrs: {
                  ...$at.parent.attrs,
                  ...($at.parent.attrs.sourceParagraph
                    ? {
                        spaceBefore: $at.parent.attrs.spaceBefore ?? '0pt',
                        spaceAfter: $at.parent.attrs.spaceAfter ?? '0pt',
                      }
                    : {}),
                  sourceParagraph: null,
                },
              },
            ]);
            at += 2;
          }
          const breakNode = state.schema.nodes[
            kind === 'page' ? 'wordPageBreak' : 'wordColumnBreak'
          ].create(null, null, marks);
          if (kind === 'page') {
            // Word gives the standalone page-break paragraph normal paragraph
            // layout, while both text portions retain their original layout.
            const font = type?.isInSet(marks)?.attrs;
            // Native insertion at the paragraph start leaves named collapsed
            // bookmarks before the break. Move the source identity with that
            // boundary; live section ends still map to the following text.
            const leadingSource = !$at.parentOffset ? $at.parent.attrs.sourceParagraph : null;
            if (leadingSource)
              tr.setNodeAttribute(at - 1, 'sourceParagraph', null)
                .setNodeAttribute(at - 1, 'spaceBefore', paragraph.spaceBefore ?? '0pt')
                .setNodeAttribute(at - 1, 'spaceAfter', paragraph.spaceAfter ?? '0pt');
            const boundary = state.schema.nodes.paragraph.create(
              {
                sourceParagraph: leadingSource,
                paragraphFontSize: font?.fontSize ?? paragraph.paragraphFontSize,
                paragraphFontFamily: font?.fontFamily ?? paragraph.paragraphFontFamily,
                paragraphLineHeight: '1',
                spaceBefore: '0pt',
                spaceAfter: '0pt',
                widowControl: true,
              },
              breakNode,
            );
            tr.insert(at - 1, boundary);
            at += boundary.nodeSize;
          } else {
            tr.insert(at, breakNode);
            at++;
          }
          tr.setSelection(TextSelection.create(tr.doc, at)).ensureMarks(marks).scrollIntoView();
          return true;
        },
    };
  },
  addKeyboardShortcuts() {
    return { 'Mod-Enter': () => this.editor.commands.insertWordFlowBreak('page') };
  },
});

/** A column boundary consumes one semantic offset, just like Word's U+000E. */
export const WordColumnBreak = WordPageBreak.extend({
  name: 'wordColumnBreak',
  parseHTML: () => [{ tag: 'span[data-word-column-break]' }],
  renderHTML: () => [
    'span',
    {
      'data-word-column-break': 'true',
      'aria-label': 'Column break',
      contenteditable: 'false',
      style: 'display:block;break-after:column',
    },
    ['br'],
  ],
  renderText: () => '\u000e',
  addCommands: () => ({}),
  addKeyboardShortcuts() {
    return { 'Mod-Shift-Enter': () => this.editor.commands.insertWordFlowBreak('column') };
  },
});
