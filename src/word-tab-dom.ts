/** ProseMirror can put a leaf node's decorations on an enclosing inline
 * fragment span. Keep the semantic tab separate from its measured paint box. */
export function wordTabPaintBox(tab: HTMLElement): HTMLElement {
  for (let current: HTMLElement | null = tab; current && current.tagName !== 'P'; current = current.parentElement)
    if (current.hasAttribute('data-word-tab-measured')) return current;
  return tab;
}
