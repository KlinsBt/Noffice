import type { Editor } from '@tiptap/core';
import type { Mark } from '@tiptap/pm/model';
import { TextSelection } from '@tiptap/pm/state';
import { closeHistory } from '@tiptap/pm/history';
export type CaseMode = 'upper' | 'lower' | 'title' | 'sentence' | 'toggle';
export const fontSteps = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
export function nextFontSize(size: number, grow: boolean) {
  if (grow && size < 8) return Math.min(8, size + 1);
  if (!grow && size <= 8) return Math.max(1, size - 1);
  if (grow && size >= 72) return Math.min(1638, (Math.floor(size / 10) + 1) * 10);
  if (!grow && size > 72) return Math.max(72, (Math.ceil(size / 10) - 1) * 10);
  return grow
    ? fontSteps.find((n) => n > size) || 72
    : [...fontSteps].reverse().find((n) => n < size) || 8;
}
/** Replace selected text runs, retaining their marks and structural/inline boundaries. */
export function changeCase(editor: Editor, mode: CaseMode) {
  const { from, to, empty } = editor.state.selection;
  if (empty) return false;
  const parts: { from: number; to: number; text: string; marks: readonly Mark[] }[] = [];
  let wordStart = true,
    sentenceStart = true,
    lastEnd = -1;
  editor.state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) return;
    const start = Math.max(from, pos),
      end = Math.min(to, pos + node.nodeSize);
    if (start === end) return;
    if (start !== lastEnd) {
      wordStart = true;
      sentenceStart = true;
    }
    const source = node.text!.slice(start - pos, end - pos);
    let text = '';
    if (mode === 'upper') text = source.toUpperCase();
    else if (mode === 'lower') text = source.toLowerCase();
    else
      for (const ch of source) {
        const letter = /\p{L}/u.test(ch);
        text +=
          mode === 'toggle'
            ? ch === ch.toUpperCase()
              ? ch.toLowerCase()
              : ch.toUpperCase()
            : letter
              ? (mode === 'title' ? wordStart : sentenceStart)
                ? ch.toUpperCase()
                : ch.toLowerCase()
              : ch;
        if (letter) {
          wordStart = false;
          sentenceStart = false;
        } else {
          if (!/[\p{N}'’]/u.test(ch)) wordStart = true;
          if (/[.!?]/.test(ch)) sentenceStart = true;
        }
      }
    parts.push({ from: start, to: end, text, marks: node.marks });
    lastEnd = end;
  });
  const tr = closeHistory(editor.state.tr);
  for (const part of parts.reverse())
    tr.replaceWith(part.from, part.to, editor.schema.text(part.text, part.marks));
  tr.setSelection(TextSelection.create(tr.doc, tr.mapping.map(from, -1), tr.mapping.map(to, 1)));
  editor.view.dispatch(tr);
  editor.view.focus();
  return true;
}
const formatMarks = new Set([
  'bold',
  'italic',
  'underline',
  'strike',
  'subscript',
  'superscript',
  'textStyle',
  'highlight',
]);
export type TextFormat = readonly Mark[];
export function captureTextFormat(editor: Editor): TextFormat {
  const { from, to, empty, $from } = editor.state.selection;
  let marks = editor.state.storedMarks || $from.marks();
  if (!empty) {
    let found = false;
    editor.state.doc.nodesBetween(from, to, (node) => {
      if (!found && node.isText) {
        marks = node.marks;
        found = true;
      }
    });
  }
  return marks.filter((mark) => formatMarks.has(mark.type.name));
}
export function applyTextFormat(editor: Editor, marks: TextFormat) {
  const { from, to, empty } = editor.state.selection;
  if (empty) return false;
  const tr = closeHistory(editor.state.tr);
  for (const name of formatMarks)
    if (editor.schema.marks[name]) tr.removeMark(from, to, editor.schema.marks[name]);
  for (const mark of marks) tr.addMark(from, to, mark);
  editor.view.dispatch(tr);
  return true;
}
/** Ordinary paragraph indentation; list nesting uses the list commands instead. */
export function indentParagraphs(editor: Editor, outdent = false) {
  if (editor.isActive('listItem'))
    return outdent
      ? editor.commands.liftListItem('listItem')
      : editor.commands.sinkListItem('listItem');
  const { from, to, $from } = editor.state.selection,
    tr = closeHistory(editor.state.tr);
  const positions = new Map<number, typeof $from.parent>();
  if ($from.parent.isTextblock) positions.set($from.before(), $from.parent);
  editor.state.doc.nodesBetween(from, to, (node, pos) => {
    if (['paragraph', 'heading'].includes(node.type.name)) positions.set(pos, node);
  });
  for (const [pos, node] of positions) {
    if (!['paragraph', 'heading'].includes(node.type.name)) continue;
    const raw = String(node.attrs.indentStart || '0pt');
    const points = (parseFloat(raw) || 0) * (raw.endsWith('px') ? 0.75 : 1);
    const value = Math.min(720, Math.max(0, points + (outdent ? -36 : 36)));
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, indentStart: `${value}pt` });
  }
  editor.view.dispatch(tr);
  editor.view.focus();
  return true;
}
