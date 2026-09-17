import type { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { TextSelection } from '@tiptap/pm/state';

/** Native navigation can finish before the selectionchange observer. Document
 * boundary shortcuts already own their model selection and restore the DOM. */
export function syncWordNavigation(editor: Editor, event: KeyboardEvent) {
  if (
    /^(ArrowLeft|ArrowRight|ArrowUp|ArrowDown|Home|End|PageUp|PageDown)$/.test(event.key) &&
    !((event.ctrlKey || event.metaKey) && /^(Home|End)$/.test(event.key))
  )
    syncWordSelection(editor);
}

export function syncWordSelection(editor: Editor) {
  // A native Shift+Arrow/End selection can precede ProseMirror's selectionchange
  // observer. Capture it before opening a dialog moves focus out of the document.
  const dom = editor.view.dom.ownerDocument.getSelection();
  if (
    editor.view.hasFocus() &&
    editor.state.selection instanceof TextSelection &&
    dom?.anchorNode &&
    dom.focusNode &&
    editor.view.dom.contains(dom.anchorNode) &&
    editor.view.dom.contains(dom.focusNode)
  ) {
    const anchor = editor.view.posAtDOM(dom.anchorNode, dom.anchorOffset),
      head = editor.view.posAtDOM(dom.focusNode, dom.focusOffset);
    if (anchor !== editor.state.selection.anchor || head !== editor.state.selection.head)
      editor.view.dispatch(
        editor.state.tr.setSelection(
          TextSelection.between(editor.state.doc.resolve(anchor), editor.state.doc.resolve(head)),
        ),
      );
  }
}

export function selectedWordLink(editor: Editor) {
  syncWordSelection(editor);
  if (editor.state.selection.empty && editor.isActive('link'))
    editor.commands.extendMarkRange('link');
  const { from, to } = editor.state.selection;
  return {
    href: editor.getAttributes('link').href || '',
    text: editor.state.doc.textBetween(from, to, ' '),
  };
}
export function applyWordLink(
  editor: Editor,
  address: string | null,
  text: string,
  originalText: string,
) {
  const href = address?.trim();
  if (
    href !== undefined &&
    (!href ||
      href.length > 8192 ||
      !/^(https?:\/\/|mailto:)/i.test(href) ||
      /[\u0000-\u0020]/.test(href))
  )
    throw Error('Enter a valid https://, http:// or mailto: address.');
  const { from, to, empty } = editor.state.selection;
  const tr = closeHistory(editor.state.tr),
    mark = editor.schema.marks.link;
  if (address === null) tr.removeMark(from, to, mark);
  else {
    const label = text || href!;
    if (label.length > 200000) throw Error('Link text exceeds the paragraph editing limit.');
    if (empty || text !== originalText) {
      const marks = (
        editor.state.doc.nodeAt(from)?.marks || editor.state.selection.$from.marks()
      ).filter((m) => m.type !== mark);
      tr.replaceSelectionWith(editor.schema.text(label, [...marks, mark.create({ href })]), false);
      tr.setStoredMarks(marks);
    } else tr.addMark(from, to, mark.create({ href }));
  }
  editor.view.dispatch(tr.scrollIntoView());
  editor.view.dispatch(closeHistory(editor.state.tr));
}
