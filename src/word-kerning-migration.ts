import { wordKerningValue } from './word-kerning';

function textNodes(root: HTMLElement): Text[] {
  const walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  while (walker.nextNode())
    if (walker.currentNode.textContent) nodes.push(walker.currentNode as Text);
  return nodes;
}

/** Earlier snapshots could not author kerning. Recover uniform inherited values
 * through text edits; mixed values require unchanged text to retain identity. */
export function hydrateParagraphKerning(current: HTMLElement, source: HTMLElement) {
  const ranges: { from: number; to: number; value: number | null }[] = [];
  let offset = 0;
  for (const text of textNodes(source)) {
    const value = wordKerningValue(
      text.parentElement?.closest('[data-word-kerning]')?.getAttribute('data-word-kerning'),
    );
    const end = offset + text.length;
    if (ranges.at(-1)?.value === value) ranges.at(-1)!.to = end;
    else ranges.push({ from: offset, to: end, value });
    offset = end;
  }
  if (!ranges.some((range) => range.value !== null)) return;
  const uniform = ranges.length === 1;
  if (!uniform && current.textContent !== source.textContent)
    throw Error(
      'This saved Word paragraph has edited text with missing mixed kerning metadata. Keep a native backup and reopen the original DOCX to recover its formatting safely.',
    );
  offset = 0;
  let index = 0;
  for (const text of textNodes(current)) {
    const end = offset + text.length;
    if (
      wordKerningValue(
        text.parentElement?.closest('[data-word-kerning]')?.getAttribute('data-word-kerning'),
      ) !== null
    ) {
      offset = end;
      continue;
    }
    const fragment = current.ownerDocument.createDocumentFragment();
    let local = offset;
    while (local < end) {
      while (!uniform && index < ranges.length - 1 && ranges[index].to <= local) index++;
      const range = ranges[index];
      const to = uniform ? end : Math.min(end, range.to);
      const part = current.ownerDocument.createTextNode(
        text.data.slice(local - offset, to - offset),
      );
      if (range.value === null) fragment.append(part);
      else {
        const span = current.ownerDocument.createElement('span');
        span.dataset.wordKerning = String(range.value);
        let parent = text.parentElement;
        while (parent) {
          if (parent.style.fontSize) {
            span.style.fontSize = parent.style.fontSize;
            break;
          }
          if (parent === current) break;
          parent = parent.parentElement;
        }
        span.append(part);
        fragment.append(span);
      }
      local = to;
    }
    text.replaceWith(fragment);
    offset = end;
  }
}

/** Paragraph marks survive even with no source text. Old empty-paragraph typing
 * had no run kerning to inherit; recover only missing values, keeping explicit
 * run overrides. Nonempty source text keeps its existing run provenance. */
export function hydrateParagraphMarkKerning(current: HTMLElement, source: HTMLElement) {
  const value = wordKerningValue(source.getAttribute('data-word-paragraph-kerning'));
  if (value === null) return;
  if (wordKerningValue(current.getAttribute('data-word-paragraph-kerning')) === null)
    current.setAttribute('data-word-paragraph-kerning', String(value));
  if (source.textContent) return;
  const synthetic = source.cloneNode(true) as HTMLElement;
  const span = synthetic.ownerDocument.createElement('span');
  span.dataset.wordKerning = String(value);
  span.textContent = 'x';
  synthetic.append(span);
  hydrateParagraphKerning(current, synthetic);
}
