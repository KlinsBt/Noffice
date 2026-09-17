import type { Node, ResolvedPos } from '@tiptap/pm/model';

/** This contract deliberately excludes nested lists and multi-paragraph items.
 * One semantic list spans its section boundaries, so subsequent editing keeps
 * numbering continuity without persisting derived restart values. */
export function isFlatWordList(node: Node) {
  return ['orderedList', 'bulletList'].includes(node.type.name) && node.childCount > 0 &&
    Array.from({ length: node.childCount }, (_, i) => node.child(i)).every(item =>
      item.type.name === 'listItem' && item.childCount === 1 &&
      item.firstChild?.type.name === 'paragraph');
}

export function inFlatWordList(position: ResolvedPos) {
  return position.depth === 3 && isFlatWordList(position.node(1)) &&
    position.node(2).type.name === 'listItem' && position.parent.type.name === 'paragraph';
}
