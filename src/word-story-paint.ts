import type { EditorView } from '@tiptap/pm/view';
import type { WordLine } from './word-line-measurements';

/** Freeze measured soft lines for story paint. A wrapping inline paint span
 * cannot become one inline-block during printing: that clips the whole story.
 * Derived clones preserve semantic text/marks and never enter editor history. */
export function wordStoryParagraphPaint(view: EditorView, from: number, lines: WordLine[]) {
  const original = view.nodeDOM(from);
  if (!(original instanceof HTMLElement)) throw Error('Missing story paragraph.');
  const paint = [...original.querySelectorAll<HTMLElement>('[data-word-baseline-paint]')];
  const wraps = paint.some(
    (run) => new Set([...run.getClientRects()].map((r) => Math.round(r.top))).size > 1,
  );
  const style = getComputedStyle(original);
  if (lines.length < 2 || (!wraps && !['center', 'right', 'justify'].includes(style.textAlign)))
    return [{ element: original.cloneNode(true) as HTMLElement, top: 0 }];
  const origins = new Map<string, number>();
  for (const run of paint) {
    const offset = view.posAtDOM(run, 0) - from - 1;
    const line = lines.find((l) => offset >= l.from && offset < l.to);
    if (line && run.dataset.wordNativeBaseline)
      origins.set(run.dataset.wordNativeBaseline, line.top);
  }
  const bounds = original.getBoundingClientRect();
  const scale = bounds.width / parseFloat(style.width);
  return lines.map((line, index) => {
    const start = view.domAtPos(from + 1 + line.from, 1);
    const end = view.domAtPos(from + 1 + line.to, -1);
    const range = document.createRange();
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
    let contents: Node = range.cloneContents();
    let ancestor: Node | null = range.commonAncestorContainer;
    if (ancestor.nodeType === Node.TEXT_NODE) ancestor = ancestor.parentNode;
    while (ancestor && ancestor !== original) {
      const wrapper = ancestor.cloneNode(false);
      wrapper.appendChild(contents);
      contents = wrapper;
      ancestor = ancestor.parentNode;
    }
    const element = original.cloneNode(false) as HTMLElement;
    element.append(contents);
    element.style.height = `${line.height}px`;
    // Each clone paints one measured visual line of the original paragraph.
    // A first/hanging indent belongs only to its first line. Justification on
    // nonfinal soft lines must survive their conversion into separate boxes.
    if (index > 0) element.style.textIndent = '0';
    if (['center', 'right'].includes(style.textAlign) && line.to > line.from && scale > 0) {
      const next = view.domAtPos(from + 2 + line.from, -1);
      const first = document.createRange();
      first.setStart(start.node, start.offset); first.setEnd(next.node, next.offset);
      const left = (first.getBoundingClientRect().left - bounds.left) / scale;
      if (!Number.isFinite(left)) throw Error('Invalid aligned story line origin.');
      // A soft line's trailing spaces hang in the measured editor. Turning it
      // into a last-line clone includes those spaces in center/right alignment.
      // Freeze its actual line origin while retaining all text, including spaces.
      element.style.textAlign = 'left';
      element.style.textIndent = `${left}px`;
    }
    // Modern line decorations already carry explicit space advances. Splitting
    // them into per-word inline-blocks rounds each width again and accumulates
    // horizontal drift in the page copy. Preserve their continuous text run.
    if (style.textAlign === 'justify' && index < lines.length - 1
      && original.dataset.wordJustification !== 'modern')
      freezeWordStorySpaces(view, from, line, element, scale);
    for (const run of element.querySelectorAll<HTMLElement>('[data-word-native-baseline]')) {
      const origin = origins.get(run.dataset.wordNativeBaseline!);
      if (origin !== undefined)
        run.dataset.wordNativeBaseline = String(Number(run.dataset.wordNativeBaseline) - origin);
    }
    return { element, top: line.top };
  });
}

/** Printing transforms baseline runs into inline-blocks, which otherwise loses
 * the justified gaps inside a wrapped run. Preserve measured word advances in
 * the disposable line copy. This is paint preservation, not Word's separate
 * compatibility-dependent line-breaking algorithm. */
function freezeWordStorySpaces(view: EditorView, from: number, line: WordLine, element: HTMLElement, scale: number) {
  const paragraph = view.state.doc.nodeAt(from);
  // Cross-run shaping needs its own cluster-aware paint. Do not introduce
  // inline-block boundaries into a mixed-mark/font paragraph here.
  if (!paragraph || paragraph.childCount !== 1 || !paragraph.firstChild?.isText || !(scale > 0)) return;
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  const texts: Text[] = [];
  while (walker.nextNode()) if (!walker.currentNode.parentElement?.closest('.ProseMirror-widget,[contenteditable=false]'))
    texts.push(walker.currentNode as Text);
  if (texts.reduce((size, text) => size + text.length, 0) !== line.to - line.from) return;
  let offset = line.from;
  for (const text of texts) {
    const fragment = document.createDocumentFragment();
    for (const match of text.data.matchAll(/[^ ]+ *| +/g)) {
      const value = match[0];
      const start = view.domAtPos(from + 1 + offset, 1), end = view.domAtPos(from + 1 + offset + value.length, -1);
      const range = document.createRange(); range.setStart(start.node, start.offset); range.setEnd(end.node, end.offset);
      let width = range.getBoundingClientRect().width / scale;
      if (offset + value.length < line.to) {
        const next = view.domAtPos(from + 2 + offset + value.length, -1);
        const following = document.createRange();
        following.setStart(end.node, end.offset); following.setEnd(next.node, next.offset);
        width = (following.getBoundingClientRect().left - range.getBoundingClientRect().left) / scale;
      }
      if (!Number.isFinite(width) || width < 0) throw Error('Invalid justified story advance.');
      const span = document.createElement('span');
      span.dataset.wordStoryWord = 'true'; span.style.cssText = `display:inline-block;width:${width}px;white-space:pre`;
      span.textContent = value; fragment.append(span); offset += value.length;
    }
    text.replaceWith(fragment);
  }
  element.style.textAlign = 'left';
  element.style.whiteSpace = 'pre';
}
