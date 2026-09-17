import type { EditorState } from '@tiptap/pm/state';

/** Formatting includes each selected paragraph once, excluding a following
 * paragraph when the range ends exactly at its content start. */
export function selectedWordParagraphs(state: Pick<EditorState, 'doc' | 'selection'>) {
  const paragraphs: { pos: number; attrs: Record<string, unknown> }[] = [];
  const seen = new Set<number>();
  for (const { $from, $to } of state.selection.ranges) {
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (!['paragraph', 'heading'].includes(node.type.name)) return;
      if ($to.pos > $from.pos && $to.pos === pos + 1) return false;
      if (!seen.has(pos)) paragraphs.push({ pos, attrs: node.attrs });
      seen.add(pos);
      return false;
    });
  }
  return paragraphs;
}
