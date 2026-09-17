import { wordLineSpacing } from './word-line-spacing';

/** Bounded regular Arial10 story paint. Native source/shaded PDFs and page
 * metafiles establish separate 600dpi origin and paragraph/descent rounding.
 * This changes the disposable paint layer, never semantic spacing or height. */
export function wordStoryGridOrigin(top: number, after: number): number {
  if (!Number.isFinite(top) || !Number.isFinite(after) || top < 0 || after < 0 || after > 801 / 15)
    return top;
  const afterTwips = Math.round(after * 15);
  const origin = Math.round(Math.round(top * 15) * 5 / 12);
  const baseline = Math.round((800 + afterTwips) * 5 / 12)
    - Math.round(afterTwips * 5 / 12) - Math.round(800 * 5 / 12);
  return (origin + baseline) * .16;
}

export function wordStoryGridEligible(element: HTMLElement, lineCount: number, variableHeight = false): boolean {
  const style = getComputedStyle(element);
  if (lineCount !== 1
    || (variableHeight ? !wordLineSpacing(element.style.lineHeight, element.dataset.wordLineRule)
      : element.dataset.wordLineHeight !== 'exact' || Math.abs(parseFloat(style.lineHeight) - 800 / 15) > .0001)
    || !/^[\x20-\x7e\t]+$/.test(element.textContent || '')
    || (element.textContent?.length || 0) > 4096
    || !['transparent', 'rgba(0, 0, 0, 0)'].includes(style.backgroundColor)
    || [style.marginTop, style.marginBottom].some(value => {
      const n = parseFloat(value); return !Number.isFinite(n) || n < 0 || n > 801 / 15 + .0001;
    })) return false;
  const runs = [...element.querySelectorAll<HTMLElement>('[data-word-native-baseline]')];
  if (!runs.length || runs.some(run => variableHeight
    ? !Number.isFinite(Number(run.dataset.wordNativeBaseline)) || Number(run.dataset.wordNativeBaseline) <= 0
    : Math.abs(Number(run.dataset.wordNativeBaseline) - 32.05 / .75) > .0001))
    return false;
  const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  let text: Node | null;
  while ((text = walker.nextNode())) {
    if (!text.textContent) continue;
    const font = getComputedStyle(text.parentElement!);
    if (font.fontFamily.replace(/["']/g, '').trim().toLowerCase() !== 'arial'
      || Math.abs(parseFloat(font.fontSize) - 10 / .75) > .0001
      || !['normal', '400'].includes(font.fontWeight) || font.fontStyle !== 'normal'
      || !['normal', '100%'].includes(font.fontStretch)) return false;
  }
  return true;
}

/** Screen stories keep their measured outer bounds; only glyph paint moves.
 * Printing uses the same origin directly in its physical fragment tree. */
export function wordStoryGridHtml(html: string, top: number): string {
  if (!html.includes('data-word-story-grid-after')) return html;
  const copy = document.createElement('div');
  copy.innerHTML = html;
  for (const p of copy.querySelectorAll<HTMLElement>('[data-word-story-grid-after]')) {
    const origin = top + Number(p.dataset.wordStoryParagraphTop);
    const local = wordStoryGridOrigin(origin, Number(p.dataset.wordStoryGridAfter)) - top;
    // Ordinary flow rounds every preceding line and collapsed margin again.
    // Paint from the measured origin, retaining its fractional layer position.
    p.style.position = 'absolute';
    p.style.left = `${Number(p.dataset.wordStoryParagraphLeft || 0)}px`;
    p.style.top = `${Math.floor(local)}px`;
    p.style.translate = `0 ${local - Math.floor(local)}px`;
    p.style.margin = '0';
  }
  return copy.innerHTML;
}
