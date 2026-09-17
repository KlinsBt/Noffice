/** Recover the previously omitted Word fallback without replacing text or explicit
 * inline formatting. Source paragraphs supply paragraph-mark metrics only. */
export function hydrateParagraphFonts(current: HTMLElement, source: HTMLElement) {
  const oldSize = current.style.fontSize;
  // Older imports already resolved explicit/default/style sizes on runs. Only
  // missing run sizes used the app default; Word's pinned fallback is 10pt.
  const walker = current.ownerDocument.createTreeWalker(current, NodeFilter.SHOW_TEXT);
  const nodes: Node[] = [...current.querySelectorAll('br')];
  while (walker.nextNode()) if (walker.currentNode.textContent) nodes.push(walker.currentNode);
  for (const node of nodes) {
    let parent = node.parentElement;
    let explicit = false;
    while (parent && parent !== current) {
      if (parent.style.fontSize) explicit = true;
      parent = parent.parentElement;
    }
    if (explicit) continue;
    const span = current.ownerDocument.createElement('span');
    span.style.fontSize = oldSize || '10pt';
    node.parentNode!.insertBefore(span, node);
    span.append(node);
  }
  if (!current.style.fontSize) current.style.fontSize = source.style.fontSize;
  if (!current.style.fontFamily) current.style.fontFamily = source.style.fontFamily;
}
