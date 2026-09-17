import type { EditorState } from '@tiptap/pm/state';
import { wordDefaults } from './word-defaults';

type FontValue = { family: string | null; size: number | null };

function fontValue(attrs: Record<string, unknown>, paragraph: Record<string, unknown>): FontValue {
  const family = attrs.fontFamily || paragraph.paragraphFontFamily || wordDefaults.fontFamily;
  const size = attrs.fontSize || paragraph.paragraphFontSize || `${wordDefaults.fontSize}pt`;
  const match = typeof size === 'string' && /^(\d+(?:\.\d+)?)(pt|px)$/.exec(size);
  const points = match ? Number(match[1]) * (match[2] === 'px' ? .75 : 1) : NaN;
  return {
    family: typeof family === 'string' ? family.replace(/^["']|["']$/g, '') : null,
    size: Number.isFinite(points) && points >= 1 && points <= 1638 ? points : null,
  };
}

/** Read formatting without changing selection or stored marks. Word aggregates
 * selected inline content independently of included paragraph marks; a mark-only
 * selection and an empty caret expose the paragraph's own formatting. */
export function wordSelectionFont(state: EditorState): FontValue {
  const { from, to, empty, $from } = state.selection;
  const type = state.schema.marks.textStyle;
  if (empty) return fontValue(type?.isInSet(state.storedMarks || $from.marks())?.attrs || {}, $from.parent.attrs);
  let content: FontValue | undefined, marks: FontValue | undefined;
  const merge = (before: FontValue | undefined, next: FontValue): FontValue => before ? {
    family: before.family === next.family ? before.family : null,
    size: before.size === next.size ? before.size : null,
  } : next;
  state.doc.nodesBetween(from, to, (paragraph, pos) => {
    if (!paragraph.isTextblock) return;
    const start = pos + 1, end = start + paragraph.content.size;
    paragraph.forEach((node, offset) => {
      if (!node.isText && !['wordTab', 'wordHyphen', 'hardBreak', 'wordPageBreak', 'wordColumnBreak'].includes(node.type.name)) return;
      const left = start + offset, right = left + node.nodeSize;
      if (left < to && right > from)
        content = merge(content, fontValue(type?.isInSet(node.marks)?.attrs || {}, paragraph.attrs));
    });
    if (to > end && from <= end) marks = merge(marks, fontValue({}, paragraph.attrs));
    return false;
  });
  return content || marks || fontValue({}, $from.parent.attrs);
}
