import type { EditorView } from '@tiptap/pm/view';
import { paragraphGraphemes, type WordLine } from './word-line-measurements';
import type { WordLineProvider } from './word-column-reflow';

/** Browser measurement for uniform, space-delimited Latin paragraphs. The
 * provider measures complete candidate lines, retaining pair-kerning context
 * and hanging spaces. Other scripts, tabs, inline objects and mixed metrics
 * keep their existing explicit fallback until their shaping contracts pass. */
export function wordColumnLineProvider(
  view: EditorView,
  from: number,
  lines: WordLine[],
): WordLineProvider | null {
  const node = view.state.doc.nodeAt(from),
    dom = view.nodeDOM(from);
  const text = node && paragraphGraphemes(node)?.text;
  if (
    !node ||
    !(dom instanceof HTMLElement) ||
    !lines.length ||
    !text ||
    text.length > 4096 ||
    !text.split('\n').every((line) => /^[A-Za-z0-9]+(?: +[A-Za-z0-9]+)* *$/.test(line)) ||
    lines.some(
      (line) =>
        Math.abs(line.height - lines[0].height) > 0.001 ||
        Math.abs((line.fitHeight ?? line.height) - (lines[0].fitHeight ?? lines[0].height)) > 0.001,
    )
  )
    return null;
  let textOnly = true;
  node.forEach((child) => {
    if (!child.isText && child.type.name !== 'hardBreak') textOnly = false;
  });
  if (!textOnly) return null;
  const walker = document.createTreeWalker(dom, NodeFilter.SHOW_TEXT);
  let style: CSSStyleDeclaration | undefined,
    signature = '';
  const descriptor = (s: CSSStyleDeclaration) =>
    [
      s.fontFamily,
      s.fontSize,
      s.fontWeight,
      s.fontStyle,
      s.fontStretch,
      s.fontVariant,
      s.fontKerning,
    ].join('|');
  while (walker.nextNode()) {
    const parent = walker.currentNode.parentElement!;
    if (!walker.currentNode.textContent || parent.closest('[contenteditable=false]')) continue;
    const current = getComputedStyle(parent);
    if (
      !['normal', '0px'].includes(current.letterSpacing) ||
      !['normal', '0px'].includes(current.wordSpacing) ||
      (current.fontStretch !== '100%' && current.fontStretch !== 'normal')
    )
      return null;
    if (signature && descriptor(current) !== signature) return null;
    signature = descriptor(current);
    style = current;
  }
  if (!style) return null;
  // Canvas font shorthand cannot express all of the editor's ligature and
  // feature settings. An isolated DOM probe uses those exact properties and
  // does not silently fall back to a default font when shorthand is rejected.
  const probe = document.createElement('span');
  probe.setAttribute('aria-hidden', 'true');
  probe.style.cssText =
    'all:initial;position:fixed;left:-100000px;top:0;visibility:hidden;pointer-events:none;white-space:pre;display:inline-block;contain:layout style paint';
  for (const property of [
    'fontFamily',
    'fontSize',
    'fontWeight',
    'fontStyle',
    'fontStretch',
    'fontVariant',
    'fontKerning',
    'fontFeatureSettings',
    'fontVariationSettings',
    'letterSpacing',
    'wordSpacing',
    'textRendering',
  ] as const)
    probe.style[property] = style[property];
  const height = lines[0].height;
  const words = [...text.matchAll(/\S+ *(?:\n)?/g)].map((match) => ({
    from: match.index!,
    to: match.index! + match[0].length,
  }));
  const starts = new Map(words.map((word, index) => [word.from, index]));
  const widths = new Map<string, number>();
  const provider: WordLineProvider = (start, width) => {
    const index = starts.get(start);
    if (index === undefined || !Number.isFinite(width) || width <= 0) return null;
    let end = start;
    for (let i = index; i < words.length; i++) {
      const candidate = text.slice(start, words[i].to).trimEnd();
      let measured = widths.get(candidate);
      if (measured === undefined) {
        if (widths.size >= 4096) return null;
        if (!probe.isConnected) document.body.append(probe);
        probe.textContent = candidate;
        measured = probe.getBoundingClientRect().width;
        widths.set(candidate, measured);
      }
      if (measured > width + 0.001) break;
      end = words[i].to;
      if (text[end - 1] === '\n') break;
    }
    return end > start
      ? {
          from: start,
          to: end,
          height,
          ...(lines[0].fitHeight === undefined ? {} : { fitHeight: lines[0].fitHeight }),
        }
      : null;
  };
  provider.dispose = () => probe.remove();
  return provider;
}
