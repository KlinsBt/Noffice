import { Extension } from '@tiptap/core';
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state';
import type { EditorView } from '@tiptap/pm/view';
import { literalWordInput } from './word-hyphen';

function replaceParagraphText(view: EditorView, from: number, to: number, text: string) {
  if (view.composing || from === to || !text || /[\r\n\t]/.test(text)) return false;
  const { state } = view, a = state.doc.resolve(from), b = state.doc.resolve(to);
  if (!(state.selection instanceof TextSelection) || a.depth !== 1 || b.depth !== 1 ||
      a.parent.type.name !== 'paragraph' || b.parent.type.name !== 'paragraph' ||
      b.parentOffset !== 0 || b.index(0) !== a.index(0) + 1) return false;
  const tr = state.tr.setSelection(TextSelection.create(state.doc, from, a.end()));
  view.dispatch(literalWordInput(tr.insertText(text), from, from + text.length).scrollIntoView());
  return true;
}

/** Native TypeText keeps a paragraph mark when a selection ends immediately
 * after it. This is text replacement, not the Delete/Backspace join command. */
export const WordParagraphInput = Extension.create({
  name: 'wordParagraphInput',
  addProseMirrorPlugins() {
    return [new Plugin({
      key: new PluginKey('wordParagraphInput'),
      props: {
        handleDOMEvents: {
          beforeinput(view, event) {
            // Chromium can turn a cross-paragraph replacement into a DOM join
            // before ProseMirror's text-input hook gets a chance to handle it.
            if (event.inputType !== 'insertText' || event.isComposing || typeof event.data !== 'string') return false;
            const { from, to } = view.state.selection;
            if (!replaceParagraphText(view, from, to, event.data)) return false;
            event.preventDefault();
            return true;
          },
        },
        handleTextInput(view, from, to, text) {
          return replaceParagraphText(view, from, to, text);
        },
      },
    })];
  },
});
