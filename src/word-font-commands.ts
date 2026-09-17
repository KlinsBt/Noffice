import { Extension, type Command } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { defaultFont } from './fonts';
import { wordDefaults } from './word-defaults';
import { nextFontSize } from './word-commands';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    wordFont: {
      setWordFontFamily: (family: string) => ReturnType;
      setWordFontSize: (size: string) => ReturnType;
      stepWordFontSize: (grow: boolean) => ReturnType;
    };
  }
}

function steppedSize(value: unknown, grow: boolean) {
  if (typeof value !== 'string') return null;
  const match = /^(\d+(?:\.\d+)?)(pt|px)$/.exec(value);
  if (!match) return null;
  const before = Number(match[1]) * (match[2] === 'px' ? .75 : 1);
  if (!Number.isFinite(before) || before < 1 || before > 1638) return null;
  const after = nextFontSize(before, grow);
  return Number.isInteger(after * 2) ? { before, after, value: after + 'pt' } : null;
}

/** Native Grow/Shrink steps each selected run. Unlike explicit size entry,
 * selecting all paragraph text does not implicitly include its mark. Plan
 * first so malformed mixed input cannot leave a partially formatted range. */
function stepWordFontSize(grow: boolean): Command {
  return ({ tr, commands, dispatch }) => {
    const { from, to, empty, $from } = tr.selection;
    const type = tr.doc.type.schema.marks.textStyle;
    if (empty) {
      const active = type.isInSet(tr.storedMarks || $from.marks());
      const size = steppedSize(active?.attrs.fontSize || $from.parent.attrs.paragraphFontSize ||
        `${wordDefaults.fontSize}pt`, grow);
      return size ? commands.setWordFontSize(size.value) : false;
    }
    const edits: { from: number; to: number; mark: ReturnType<typeof type.create> }[] = [];
    const paragraphs: { pos: number; size: string }[] = [];
    let valid = true;
    tr.doc.nodesBetween(from, to, (paragraph, pos) => {
      if (!paragraph.isTextblock) return;
      const start = pos + 1, end = start + paragraph.content.size;
      const inherited = paragraph.attrs.paragraphFontSize || `${wordDefaults.fontSize}pt`;
      const affectsMark = paragraph.type.name === 'paragraph' && to > end;
      const markSize = affectsMark ? steppedSize(inherited, grow) : null;
      if (affectsMark && !markSize) { valid = false; return false; }
      const changesMark = !!markSize && markSize.before !== markSize.after;
      paragraph.forEach((node, offset) => {
        if (!node.isText && !['wordTab', 'wordHyphen', 'hardBreak', 'wordPageBreak', 'wordColumnBreak'].includes(node.type.name)) return;
        const left = start + offset, right = left + node.nodeSize;
        const selectedFrom = Math.max(from, left), selectedTo = Math.min(to, right);
        const old = type.isInSet(node.marks);
        if (selectedFrom < selectedTo) {
          const size = steppedSize(old?.attrs.fontSize || inherited, grow);
          if (!size) valid = false;
          else if (size.before !== size.after) edits.push({ from: selectedFrom, to: selectedTo,
            mark: type.create({ ...old?.attrs, fontSize: size.value }) });
        }
        if (changesMark && !old?.attrs.fontSize) {
          const mark = type.create({ ...old?.attrs, fontSize: inherited });
          if (left < from) edits.push({ from: left, to: Math.min(right, from), mark });
          if (right > to) edits.push({ from: Math.max(left, to), to: right, mark });
        }
      });
      if (changesMark) paragraphs.push({ pos, size: markSize!.value });
      return false;
    });
    if (!valid) return false;
    if (!dispatch || (!edits.length && !paragraphs.length)) return true;
    closeHistory(tr);
    for (const edit of edits) tr.addMark(edit.from, edit.to, edit.mark);
    for (const paragraph of paragraphs) tr.setNodeAttribute(paragraph.pos, 'paragraphFontSize', paragraph.size);
    return true;
  };
}

/** A paragraph mark is persistent formatting, including at an empty/end caret.
 * Its CSS also supplies text inheritance, so preserve unselected runs before
 * changing that inheritance. Text and mark edits share one history event. */
function setWordFont(key: 'fontFamily' | 'fontSize', value: string): Command {
  return ({ tr, commands, dispatch }) => {
    if (
      key === 'fontSize'
        ? !/^\d+(?:\.\d+)?pt$/.test(value) ||
          parseFloat(value) < 1 ||
          parseFloat(value) > 1638 ||
          !Number.isInteger(parseFloat(value) * 2)
        : !value.trim() || value.length > 200 || /[\x00-\x1f]/.test(value)
    )
      return false;
    if (!dispatch) return commands.setMark('textStyle', { [key]: value });
    const { from, to, empty } = tr.selection;
    const type = tr.doc.type.schema.marks.textStyle;
    const attr = key === 'fontSize' ? 'paragraphFontSize' : 'paragraphFontFamily';
    closeHistory(tr);
    if (empty && tr.selection.$from.parent.type.name === 'paragraph') {
      const cursor = tr.selection.$from;
      const text = cursor.parent.textBetween(0, cursor.parent.content.size, '\ufffc', '\ufffc');
      const offset = cursor.parentOffset;
      // Word formats an ASCII word at an interior caret without selecting it
      // or changing the paragraph mark. Boundaries remain insertion formatting.
      if (/[a-z]/i.test(text[offset - 1] || '') && /[a-z]/i.test(text[offset] || '')) {
        let left = offset,
          right = offset;
        while (left > 0 && /[a-z]/i.test(text[left - 1])) left--;
        while (right < text.length && /[a-z]/i.test(text[right])) right++;
        const start = cursor.start() + left,
          end = cursor.start() + right;
        tr.doc.nodesBetween(start, end, (node, pos) => {
          if (node.isText)
            tr.addMark(
              Math.max(start, pos),
              Math.min(end, pos + node.nodeSize),
              type.create({ ...type.isInSet(node.marks)?.attrs, [key]: value }),
            );
        });
      }
    }
    tr.doc.nodesBetween(from, to, (paragraph, pos) => {
      if (paragraph.type.name !== 'paragraph') return;
      const start = pos + 1;
      const end = start + paragraph.content.size;
      const affectsMark = empty ? from === end : to > end || (from <= start && to >= end);
      if (!affectsMark) return false;
      const normalized = value.replace(/^(["'])(.*)\1$/, '$2');
      // Attribute steps are recorded even when their values are identical.
      // Keep repeated commits out of history; selected runs are still handled
      // by setMark below, including explicit overrides of inherited formatting.
      if (paragraph.attrs[attr] === normalized) return false;
      const previous =
        paragraph.attrs[attr] || (key === 'fontSize' ? `${wordDefaults.fontSize}pt` : defaultFont);
      paragraph.forEach((node, offset) => {
        if (
          !node.isText &&
          !['wordTab', 'wordHyphen', 'hardBreak', 'wordPageBreak', 'wordColumnBreak'].includes(node.type.name)
        )
          return;
        const old = type.isInSet(node.marks);
        if (old?.attrs[key]) return;
        const left = start + offset;
        const right = left + node.nodeSize;
        const mark = type.create({ ...old?.attrs, [key]: previous });
        if (left < from) tr.addMark(left, Math.min(right, from), mark);
        if (right > to) tr.addMark(Math.max(left, to), right, mark);
      });
      tr.setNodeAttribute(pos, attr, normalized);
      return false;
    });
    if (empty) return commands.setMark('textStyle', { [key]: value });
    // Visit inline content only. Applying a fresh mark to the containing
    // paragraph first strips existing attributes and then restores them run by
    // run, creating history steps even for an unchanged value.
    for (const range of tr.selection.ranges) {
      const left = range.$from.pos, right = range.$to.pos;
      tr.doc.nodesBetween(left, right, (node, pos) => {
        if (!node.isInline) return;
        const old = type.isInSet(node.marks);
        const mark = type.create({ ...old?.attrs, [key]: value });
        if (!old?.eq(mark)) tr.addMark(Math.max(left, pos), Math.min(right, pos + node.nodeSize), mark);
      });
    }
    return true;
  };
}

export const WordFontCommands = Extension.create({
  name: 'wordFontCommands',
  addCommands() {
    return {
      setWordFontFamily: (family) => setWordFont('fontFamily', family),
      setWordFontSize: (size) => setWordFont('fontSize', size),
      stepWordFontSize,
    };
  },
});
