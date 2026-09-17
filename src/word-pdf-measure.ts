import type { EditorView } from '@tiptap/pm/view';
import type { SurfaceInput } from './word-section-surfaces';
import type { WordFragmentPlan } from './word-fragment-plan';
import {
  wordFragmentPrint,
  alignWordPrintGlyphs,
  wordPrintGlyphBaseline,
} from './word-fragment-print';
import { wordBrowserFontMetrics } from './word-browser-font-metrics';
import { validateWordPdfSnapshot, wordPdfCharacterSupported, type WordPdfSnapshot, type WordPdfDecorationSpan } from './word-pdf-model';
import { wordTabStops } from './word-tab-stops';
import { wordTabBarWidth } from './word-tab-bars';
import { wordTabPaintBox } from './word-tab-dom';
import { wordFontFeaturesValue } from './word-font-features';
import { wordTabLeaderPrinterStep } from './word-tab-leader-layout';
import { wordHyphenText } from './word-hyphen';

function pdfRunDecorations(parent: HTMLElement, box: Element) {
  let marks = 0;
  for (let ancestor: HTMLElement | null = parent; ancestor && ancestor !== box; ancestor = ancestor.parentElement) {
    const decoration = getComputedStyle(ancestor);
    const lines = decoration.textDecorationLine.split(' ');
    if (lines.some((line) => !['none', 'underline', 'line-through'].includes(line))
      || (decoration.textDecorationLine !== 'none' && (decoration.textDecorationStyle !== 'solid'
        || decoration.textDecorationColor !== decoration.color || decoration.textDecorationThickness !== 'auto'
        || decoration.textUnderlineOffset !== 'auto')) || decoration.fontVariantCaps !== 'normal'
      || decoration.textShadow !== 'none' || decoration.backgroundColor !== 'rgba(0, 0, 0, 0)'
      || decoration.opacity !== '1' || ['MARK', 'A'].includes(ancestor.tagName)
      || (['SUB', 'SUP'].includes(ancestor.tagName)
        && parent.closest<HTMLElement>('[data-word-script-paint]')?.dataset.wordScriptPaint
          !== (ancestor.tagName === 'SUP' ? 'superscript' : 'subscript')))
      throw Error('PDF export does not yet support this text formatting. Use Print / Save as PDF.');
    if (lines.includes('underline')) marks |= 1;
    if (lines.includes('line-through')) marks |= 2;
  }
  return marks;
}

function pdfRunStyle(parent: HTMLElement, box: Element, sizeOverride?: number) {
  const style = getComputedStyle(parent), size = sizeOverride ?? parseFloat(style.fontSize);
  const marks = pdfRunDecorations(parent, box);
  if (!Number.isFinite(size) || size <= 0
    || !['400', 'normal', '700', 'bold'].includes(style.fontWeight)
    || !['normal', 'italic'].includes(style.fontStyle)
    || !['normal', '100%'].includes(style.fontStretch)
    || (style.letterSpacing !== 'normal' && parseFloat(style.letterSpacing) !== 0))
    throw Error('PDF export does not yet support this font style. Use Print / Save as PDF.');
  const family = style.fontFamily.split(',')[0].trim().replace(/^['"]|['"]$/g, '');
  const rgb = /^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(1(?:\.0+)?))?\)$/.exec(style.color);
  if (!rgb) throw Error('PDF export does not yet support this text color.');
  const color = rgb.slice(1, 4).map((n) => Number(n) / 255) as [number, number, number];
  const bold = ['700', 'bold'].includes(style.fontWeight), italic = style.fontStyle === 'italic';
  const face: WordPdfSnapshot['pages'][number]['glyphs'][number]['face'] = bold
    ? italic ? 'boldItalic' : 'bold' : italic ? 'italic' : undefined;
  if ((face || marks) && !['arial', 'calibri'].includes(family.toLowerCase()))
    throw Error('PDF export does not yet support this font style. Use Print / Save as PDF.');
  return { style, size, family, color, face, marks };
}

/** Read the shared fragment layout in a temporary, hidden tree. PDF drawing
 * consumes these positions instead of repaginating with another font engine. */
export function measureWordPdf(
  view: EditorView,
  input: SurfaceInput,
  plan: WordFragmentPlan,
  content: string,
): WordPdfSnapshot {
  const tree = wordFragmentPrint(view, input, plan);
  Object.assign(tree.style, {
    position: 'absolute',
    left: '0',
    top: '0',
    visibility: 'hidden',
    pointerEvents: 'none',
    display: 'block',
  });
  for (const page of tree.children) (page as HTMLElement).style.position = 'relative';
  view.dom.after(tree);
  try {
    if (tree.querySelector('[data-word-tab-layout-unsupported="true"]'))
      throw Error('PDF export cannot lay out the document\'s unsupported tab definitions. The original DOCX is preserved.');
    if ([...tree.querySelectorAll<HTMLElement>('[data-word-tabs]')].some((paragraph) =>
      wordTabStops(paragraph.dataset.wordTabs)?.some((stop) => stop.alignment === 'bar') &&
      paragraph.dataset.wordBarPaint !== 'true'))
      throw Error('PDF export needs measured bar tab stops. Export DOCX to preserve them.');
    alignWordPrintGlyphs(tree);
    tree.style.display = 'block';
    const cache = new Map<string, { ascent: number; rangeAscent: number }>();
    const pages: WordPdfSnapshot['pages'] = [];
    let count = 0, leaderCount = 0, run = 0;
    let bodyText = '';
    for (const [index, box] of [...tree.children].entries()) {
      const physical = plan.pages[index],
        bounds = box.getBoundingClientRect();
      const scale = bounds.width / physical.width;
      if (!(scale > 0 && Number.isFinite(scale)))
        throw Error('The page layout is not ready for PDF export.');
      const glyphs: WordPdfSnapshot['pages'][number]['glyphs'] = [];
      const decorations: WordPdfDecorationSpan[] = [];
      let decorationRun = 0, decorationParagraph: Element | null = null;
      const decorate = (parent: HTMLElement, marks: number,
        span: Omit<WordPdfDecorationSpan, 'run' | 'decoration'>) => {
        const paragraph = parent.closest('p,h1,h2,h3,h4,h5,h6');
        if (!marks || paragraph !== decorationParagraph) decorationRun++;
        decorationParagraph = paragraph;
        if (marks) decorations.push({ ...span, decoration: marks, run: decorationRun });
      };
      const storyText = new Map<string, string>();
      let group: { paragraph: Element | null; signature: string } | undefined;
      const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        const text = walker.currentNode as Text,
          parent = text.parentElement!;
        const semanticTab = parent.closest<HTMLElement>('[data-word-tab]');
        const tab = semanticTab && wordTabPaintBox(semanticTab);
        if (tab) {
          group = undefined;
          if (text.data !== '\t' || tab.dataset.wordTabMeasured !== 'true')
            throw Error('PDF export needs measured tab stops. Use Print / Save as PDF.');
          // The semantic leaf can be inside ProseMirror's paint wrapper and
          // carry its own marks. Resolve its typography, not the wrapper's.
          if (pdfRunDecorations(parent, box)) {
            const tabStyle = pdfRunStyle(parent, box);
            const baseline = wordPrintGlyphBaseline(parent);
            if (baseline === null)
              throw Error('PDF export needs a measured baseline for marked tabs. Use Print / Save as PDF.');
            const rect = tab.getBoundingClientRect();
            decorate(parent, tabStyle.marks, { x: (rect.left - bounds.left) / scale * .75,
              y: baseline * .75, width: rect.width / scale * .75, size: tabStyle.size * .75,
              family: tabStyle.family, color: tabStyle.color, ...(tabStyle.face ? { face: tabStyle.face } : {}) });
          } else decorationRun++;
          if (tab.dataset.wordTabLeader !== 'none') {
            const leader = ({ dot: '.', hyphen: '-', underscore: '_', heavy: '_', middleDot: '\u00b7' } as Record<string, string>)[tab.dataset.wordTabLeader || ''];
            const offsets: unknown = JSON.parse(tab.dataset.wordTabLeaderOffsets || 'null');
            const baseline = wordPrintGlyphBaseline(tab);
            if (!leader || tab.dataset.wordTabLeaderPainted !== 'true' || !Array.isArray(offsets)
              || offsets.length > 5000 || (leaderCount += offsets.length) > 5000 || baseline === null
              || tab.dataset.wordTabLeaderGlyph !== (offsets.length ? leader : ''))
              throw Error('PDF export needs measured leader glyphs and baselines. Use Print / Save as PDF.');
            const { size, family, color, face } = pdfRunStyle(tab, box,
              parseFloat(getComputedStyle(tab).getPropertyValue('--word-tab-leader-size')));
            const rect = tab.getBoundingClientRect();
            for (const [i, offset] of offsets.entries()) {
              // Rounding the advance and then its physical origin can place
              // the first cell before the unrounded DOM edge by at most two
              // half printer steps. This is measured paint, not negative text.
              if (typeof offset !== 'number' || !Number.isFinite(offset) || offset < -wordTabLeaderPrinterStep / .75
                || offset >= rect.width / scale || (i > 0 && offset <= offsets[i - 1]) || ++count > 50000)
                throw Error('PDF export needs valid leader cell positions. Use Print / Save as PDF.');
              glyphs.push({ text: leader, x: ((rect.left - bounds.left) / scale + offset) * .75,
                y: baseline * .75, size: size * .75, family, color, ...(face ? { face } : {}) });
            }
          }
          const storyElement = parent.closest<HTMLElement>('[data-word-story]');
          const story = storyElement?.dataset.wordStoryRenderKey || storyElement?.dataset.wordStory;
          if (story) storyText.set(story, (storyText.get(story) || '') + '\t');
          else bodyText += '\t';
          continue;
        }
        const hyphen = parent.closest<HTMLElement>('[data-word-hyphen]');
        const hyphenKind = hyphen?.dataset.wordHyphen;
        if (hyphen && (wordHyphenText(hyphenKind) === null || text.data !== wordHyphenText(hyphenKind)))
          throw Error('PDF export needs a valid Word hyphen character.');
        if (!text.data || parent.closest('.ProseMirror-widget') ||
          (parent.closest('[contenteditable=false]') && parent.closest('[contenteditable=false]') !== hyphen)) {
          group = undefined;
          decorationRun++;
          continue;
        }
        const { style, size, family, color, face, marks } = pdfRunStyle(parent, box);
        const listMarker = parent.hasAttribute('data-word-list-marker-paint');
        const features = wordFontFeaturesValue(parent.closest('[data-word-font-features]')
          ?.getAttribute('data-word-font-features')) ?? wordFontFeaturesValue(parent.closest('[data-word-paragraph-font-features]')
          ?.getAttribute('data-word-paragraph-font-features')) ?? 0;
        const paragraph = parent.closest('p,h1,h2,h3,h4,h5,h6');
        const signature = JSON.stringify([family, face, size, features, style.fontKerning,
          style.fontVariantLigatures, style.fontFeatureSettings, style.fontVariationSettings,
          style.fontOpticalSizing, parent.closest('[lang]')?.getAttribute('lang')]);
        // Edit-session spans are provenance, not shaping boundaries. Native
        // Word also shapes a compatible cluster across different run colors;
        // its first character supplies the cluster's painted color.
        if (listMarker || !paragraph || !group || group.paragraph !== paragraph || group.signature !== signature) run++;
        group = { paragraph, signature };
        const metrics = wordBrowserFontMetrics(
          style,
          size,
          parseFloat(style.lineHeight) || size * 1.2,
          scale,
          cache,
        );
        const semanticBaseline = wordPrintGlyphBaseline(parent);
        const wrappedBaselines = new Map<number, number | null>();
        for (let offset = 0; offset < text.length; offset++) {
          if (++count > 50000 || (!hyphen && !wordPdfCharacterSupported(text.data[offset]) && !(listMarker && text.data[offset] === '\u2022')))
            throw Error('This text cannot yet be exported to PDF accurately.');
          const storyElement = parent.closest<HTMLElement>('[data-word-story]');
          const story = storyElement?.dataset.wordStoryRenderKey || storyElement?.dataset.wordStory;
          if (story) storyText.set(story, (storyText.get(story) || '') + text.data[offset]);
          else if (!listMarker) bodyText += text.data[offset];
          const range = document.createRange();
          range.setStart(text, offset);
          range.setEnd(text, offset + 1);
          const rect = range.getBoundingClientRect();
          if (semanticBaseline === null && !wrappedBaselines.has(rect.top))
            wrappedBaselines.set(rect.top, wordPrintGlyphBaseline(parent, rect));
          const baseline = semanticBaseline ?? wrappedBaselines.get(rect.top);
          // An optional control contributes semantic text even when no glyph
          // is painted. A used break paints the same native hyphen-minus.
          if (hyphenKind === 'optional' && rect.width === 0) continue;
          const glyph = {
            text: hyphenKind === 'literal' ? '\u00ad' : hyphen ? '-' : text.data[offset],
            ...(hyphenKind === 'literal' ? { literalHyphen: true as const } : {}),
            ...(listMarker ? { listMarker: true as const } : {}),
            x: ((rect.left - bounds.left) / scale) * 0.75,
            y: (baseline ?? (rect.top - bounds.top) / scale + metrics.rangeAscent) * 0.75,
            size: size * 0.75,
            family,
            color,
            features,
            decoration: marks,
            run,
            ...(face ? { face } : {}),
          };
          glyphs.push(glyph);
          decorate(parent, marks, { ...glyph, width: rect.width / scale * .75 });
        }
        if (listMarker) group = undefined;
      }
      const rules = (plan.separators || [])
        .filter((rule) => rule.page === index)
        .map((rule) => ({
          x: (rule.left - physical.left - 0.5) * 0.75,
          y: (rule.top - physical.top) * 0.75,
          width: 0.75,
          height: rule.height * 0.75,
        }));
      for (const paragraph of box.querySelectorAll<HTMLElement>('[data-word-bar-paint="true"]')) {
        const rect = paragraph.getBoundingClientRect();
        const offsets: number[] = JSON.parse(paragraph.dataset.wordBarOffsets || '[]');
        for (const offset of offsets) rules.push({
          x: ((rect.left - bounds.left) / scale + offset) * 0.75,
          y: ((rect.top - bounds.top) / scale) * 0.75,
          width: wordTabBarWidth * 0.75,
          height: rect.height / scale * 0.75,
        });
      }
      for (const story of physical.stories || [])
        if (
          (storyText.get(story.renderKey || story.path) || '') !==
          input.storyText?.[story.renderKey || story.path]
        )
          throw Error('The PDF layout does not contain the complete header or footer text.');
      pages.push({ width: physical.width * 0.75, height: physical.height * 0.75, glyphs, rules, decorations });
    }
    const snapshot = { content, pages };
    validateWordPdfSnapshot(snapshot);
    if (bodyText !== view.state.doc.textBetween(0, view.state.doc.content.size, '',
      (leaf) => leaf.type.name === 'wordTab' ? '\t' : leaf.type.name === 'wordHyphen' ? wordHyphenText(leaf.attrs.kind) || '' : ''))
      throw Error('The PDF layout does not contain the complete document text.');
    return snapshot;
  } finally {
    tree.remove();
  }
}
