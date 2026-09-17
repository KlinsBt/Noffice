import type { EditorView } from '@tiptap/pm/view';
import type { WordFragmentPlan } from './word-fragment-plan';
import { sectionPrintPage, type SurfaceInput } from './word-section-surfaces';
import { wordBrowserFontMetrics } from './word-browser-font-metrics';
import { wordStoryGridOrigin } from './word-story-grid';
import { wordBodyGridOffsets } from './word-body-grid';
import { wordTabPaintBox } from './word-tab-dom';

/** A temporary print tree contains each semantic fragment exactly once. It never
 * enters the editable document, storage or DOCX. Slicing the actual DOM retains
 * font decorations and all enclosing marks without clipping duplicated text. */
export function wordFragmentPrint(
  view: EditorView,
  input: SurfaceInput,
  plan: WordFragmentPlan,
): HTMLElement {
  const root = document.createElement('div');
  root.className = 'document-page word-fragment-print';
  root.setAttribute('aria-hidden', 'true');
  const pages = plan.pages.map((page) => {
    const section = input.sections.findIndex((s) => s.id === page.sectionId);
    const box = document.createElement('div');
    box.className = 'word-fragment-print-page';
    box.style.cssText =
      `page:nofficeFragment${sectionPrintPage(input, section)};` +
      `width:${page.width}px;height:${page.height}px;overflow:hidden`;
    root.append(box);
    return { box, page };
  });
  const firstOrigins = new Map<number, number>();
  const bodyGrid = wordBodyGridOffsets(view, input, plan);
  for (const separator of plan.separators || []) {
    const { box, page } = pages[separator.page];
    const rule = document.createElement('div');
    rule.className = 'word-column-separator';
    rule.style.cssText =
      `position:absolute;border-left:1px solid #000;width:0;left:0;` +
      `transform:translateX(${separator.left - page.left - 0.5}px);` +
      `top:${separator.top - page.top}px;height:${separator.height}px`;
    box.append(rule);
  }
  for (const fragment of plan.fragments) {
    const b = input.blocks[fragment.block],
      original = view.nodeDOM(b.from);
    if (!(original instanceof HTMLElement)) throw Error('Missing printable paragraph.');
    const paragraph = original.cloneNode(false) as HTMLElement;
    // Empty terminal lines can reanchor the editable paragraph on a later
    // page. Normalize against the first fragment's unshifted origin, so that
    // anchor displacement never enters the printed paragraph's local baseline.
    if (!firstOrigins.has(fragment.block))
      firstOrigins.set(fragment.block, fragment.top - fragment.shift);
    paragraph.dataset.wordFragmentSourceTop = String(
      fragment.top - fragment.shift - firstOrigins.get(fragment.block)!,
    );
    paragraph.dataset.wordFragmentTop = String(fragment.top - plan.pages[fragment.page].top + (bodyGrid.get(fragment.block) || 0));
    // The screen paragraph's paint translation is already included in the
    // physical origin below. It must not be applied a second time to print.
    paragraph.style.translate = '';
    paragraph.removeAttribute('data-word-body-grid');
    paragraph.removeAttribute('id');
    paragraph.removeAttribute('contenteditable');
    paragraph.classList.remove('word-surface-block', 'word-measured-block');
    const start = view.domAtPos(b.from + 1 + fragment.from, 1);
    const end = view.domAtPos(b.from + 1 + fragment.to, -1);
    const range = document.createRange();
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
    let contents: Node = range.cloneContents();
    let ancestor: Node | null = range.commonAncestorContainer;
    if (ancestor.nodeType === Node.TEXT_NODE) ancestor = ancestor.parentNode;
    // cloneContents omits ancestors of its common container; their font and
    // baseline styles are still required when a fragment fits within one run.
    while (ancestor && ancestor !== original) {
      const wrapper = ancestor.cloneNode(false);
      wrapper.appendChild(contents);
      contents = wrapper;
      ancestor = ancestor.parentNode;
    }
    paragraph.append(contents);
    paragraph.removeAttribute('data-word-list-marker');
    let listMarker: HTMLElement | undefined;
    if (b.list?.marker && fragment.from === 0) {
      const marker = document.createElement('span');
      marker.dataset.wordListMarkerPaint = 'true';
      marker.textContent = b.list.marker;
      Object.assign(marker.style, {
        position: 'relative',
        fontFamily: b.list.definition.font, fontSize: `${b.list.definition.size}pt`,
        fontWeight: 'normal', fontStyle: 'normal', color: 'black', whiteSpace: 'pre',
        transform: 'translateY(var(--word-list-marker-shift,0px))',
        lineHeight: 'var(--word-list-marker-height,normal)',
      });
      const baseline = original.dataset.wordListBaseline;
      if (baseline !== undefined) {
        marker.dataset.wordNativeBaseline = baseline;
        marker.dataset.wordBaselinePaint = 'true';
        marker.style.setProperty('--word-baseline-shift', getComputedStyle(original).getPropertyValue('--word-list-marker-shift'));
      }
      listMarker = marker;
    }
    // Page displacement belongs to the editable canvas. Each print paragraph
    // already has its own physical page/column position.
    for (const span of paragraph.querySelectorAll('.word-page-fragment'))
      span.classList.remove('word-page-fragment');
    // A side:-1 widget at a semantic boundary is excluded from the range's
    // start but included at its end. Keep the strut belonging to this fragment,
    // not the next page's line. The leading strut supplies minimum-spacing
    // baseline placement even though it contains no document characters.
    for (const strut of paragraph.querySelectorAll<HTMLElement>('[data-word-line-strut]')) {
      const from = Number(strut.dataset.wordLineFrom);
      if (from < fragment.from || (from >= fragment.to && from !== fragment.from)) strut.remove();
    }
    const leading = original.querySelector<HTMLElement>(
      `[data-word-line-strut][data-word-line-from="${fragment.from}"]`,
    );
    if (
      leading &&
      !paragraph.querySelector(`[data-word-line-strut][data-word-line-from="${fragment.from}"]`)
    )
      paragraph.prepend(leading.cloneNode(true));
    const { box, page } = pages[fragment.page];
    // Screen paragraphs retain their leading CSS margin. Print clears margins,
    // so include that physical offset exactly once in the fragment origin.
    const left = fragment.left - page.left + (b.indentStart || 0), top = fragment.top - page.top + (bodyGrid.get(fragment.block) || 0);
    if (fragment.from > 0) paragraph.style.textIndent = '0';
    Object.assign(paragraph.style, {
      position: 'absolute',
      // As for stories below, retain fractional physical origins in a layer
      // transform instead of Chromium's rounded positioned-text origin.
      left: `${Math.floor(left)}px`,
      top: `${Math.floor(top)}px`,
      transform: `translate(${left - Math.floor(left)}px,${top - Math.floor(top)}px)`,
      width: `${input.sections[b.section].columns ? input.sections[b.section].columns!.widths[fragment.column] / 15 : b.width}px`,
      height: `${fragment.height}px`,
      margin: '0',
      page: 'auto',
      breakBefore: 'auto',
      breakAfter: 'auto',
    });
    if (listMarker && b.list) {
      // Positioned sibling paragraphs preserve Word's marker-before-text PDF
      // order. An absolute child of the text paragraph paints after its text.
      const markerParagraph = document.createElement('p');
      markerParagraph.dataset.wordFragmentSourceTop = paragraph.dataset.wordFragmentSourceTop;
      markerParagraph.dataset.wordFragmentTop = paragraph.dataset.wordFragmentTop;
      Object.assign(markerParagraph.style, {
        position: 'absolute', left: `${Math.floor(left - b.list.definition.hanging / 15)}px`,
        top: paragraph.style.top, transform: `translate(${left - b.list.definition.hanging / 15 - Math.floor(left - b.list.definition.hanging / 15)}px,${top - Math.floor(top)}px)`,
        width: `${b.list.definition.hanging / 15}px`, height: paragraph.style.height,
        margin: '0', padding: '0', textIndent: '0', fontSize: `${b.list.definition.size}pt`,
        fontFamily: b.list.definition.font, lineHeight: getComputedStyle(original).getPropertyValue('--word-list-marker-height') || 'normal',
      });
      listMarker.style.transform = `translateY(${getComputedStyle(original).getPropertyValue('--word-list-marker-shift') || '0px'})`;
      listMarker.style.lineHeight = 'inherit';
      markerParagraph.append(listMarker);
      box.append(markerParagraph);
    }
    box.append(paragraph);
  }
  for (const { box, page } of pages) {
    // Native Word emits both side stories before the main story in its PDF
    // content stream. Keep that order as well as their physical page positions.
    const bodyStart = box.firstChild;
    for (const story of page.stories || []) {
      const html = input.storyHtml?.[story.renderKey || story.path];
      if (html === undefined) throw Error('A printable header or footer is missing.');
      const wrapper = document.createElement('div');
      wrapper.innerHTML = html;
      for (const paragraph of wrapper.querySelectorAll<HTMLElement>(
        '[data-word-story-paragraph-top]',
      )) {
        paragraph.dataset.wordFragmentSourceTop = '0';
        const origin = story.top + Number(paragraph.dataset.wordStoryParagraphTop);
        paragraph.dataset.wordFragmentTop = String(paragraph.hasAttribute('data-word-story-grid-after')
          ? wordStoryGridOrigin(origin, Number(paragraph.dataset.wordStoryGridAfter)) : origin);
        paragraph.dataset.wordStory = story.path;
        if (story.renderKey) paragraph.dataset.wordStoryRenderKey = story.renderKey;
        const top = Number(paragraph.dataset.wordFragmentTop);
        Object.assign(paragraph.style, {
          position: 'absolute',
          left: `${story.left + Number(paragraph.dataset.wordStoryParagraphLeft || 0)}px`,
          // Chromium rounds positioned text origins before applying a layer
          // transform. Preserve a fractional footer origin in that transform.
          top: `${Math.floor(top)}px`,
          transform: `translateY(${top - Math.floor(top)}px)`,
          width: `${Number(paragraph.dataset.wordStoryParagraphWidth || story.width)}px`,
          margin: '0',
          page: 'auto',
          breakBefore: 'auto',
          breakAfter: 'auto',
        });
        for (const run of paragraph.querySelectorAll<HTMLElement>('[data-word-baseline-paint]')) {
          if (run.style.transform) continue;
          const shift = parseFloat(run.style.getPropertyValue('--word-baseline-shift'));
          if (!Number.isFinite(shift)) continue;
          // The measured unwrapped single-line path also needs a transformed
          // paint run to retain its fractional baseline in Chromium's PDF.
          Object.assign(run.style, {
            display: 'inline-block',
            whiteSpace: 'pre',
            top: '0px',
            transform: `translateY(${shift}px)`,
          });
        }
        box.insertBefore(paragraph, bodyStart);
      }
    }
  }
  return root;
}

/** PDF drawing can consume the exact ephemeral baseline without measuring
 * it again through the browser's fractional CSS layout/transform grid. */
export function wordPrintGlyphBaseline(parent: HTMLElement, glyph?: DOMRect): number | null {
  const run = parent.closest<HTMLElement>('[data-word-native-baseline]');
  const paragraph = parent.closest<HTMLElement>('[data-word-fragment-source-top]');
  if (!run || !paragraph || !paragraph.contains(run)) return null;
  // The tag anchors the first line of an inline paint span. A span can wrap;
  // its later glyphs retain the measured displacement from that first line.
  // Falling back to rounded glyph bounds only for a wrapped run otherwise
  // splits one physical baseline at an adjacent short edit-session span.
  const rectangles = [...run.getClientRects()];
  let offset = 0;
  if (rectangles.some(rectangle => rectangle.top !== rectangles[0]?.top)) {
    if (!glyph || !rectangles.some(r => r.top === glyph.top && r.height === glyph.height)) return null;
    const scale = paragraph.getBoundingClientRect().width / parseFloat(paragraph.style.width);
    if (!(scale > 0 && Number.isFinite(scale))) return null;
    offset = (glyph.top - rectangles[0].top) / scale;
  }
  const baseline = Number(run.dataset.wordNativeBaseline);
  const sourceTop = Number(paragraph.dataset.wordFragmentSourceTop);
  const top = Number(paragraph.dataset.wordFragmentTop);
  return [baseline, sourceTop, top].every(Number.isFinite) ? top + baseline - sourceTop + offset : null;
}

/** Slicing a continuation line changes its local paint origin. Rebase fractional
 * glyph paint after attaching the clone, under the active print font/zoom grid. */
export function alignWordPrintGlyphs(root: HTMLElement) {
  root.style.display = 'block';
  const cache = new Map<string, { ascent: number; rangeAscent: number }>();
  try {
    for (const paragraph of root.querySelectorAll<HTMLElement>('p')) {
      const bounds = paragraph.getBoundingClientRect();
      const scale = bounds.width / parseFloat(paragraph.style.width);
      if (!(scale > 0 && Number.isFinite(scale))) continue;
      for (const span of paragraph.querySelectorAll<HTMLElement>(
        '[data-word-baseline-paint],[data-word-run-leading]',
      )) {
        let shift = parseFloat(span.style.getPropertyValue('--word-baseline-shift'));
        if (!Number.isFinite(shift)) continue;
        const baseline = Number(span.dataset.wordNativeBaseline);
        const sourceTop = Number(paragraph.dataset.wordFragmentSourceTop);
        if (Number.isFinite(baseline) && Number.isFinite(sourceTop)) {
          const walker = document.createTreeWalker(span, NodeFilter.SHOW_TEXT);
          const text = walker.nextNode();
          if (text?.textContent) {
            const semanticTab = span.textContent === '\t' && text.parentElement?.closest<HTMLElement>('[data-word-tab]');
            const tab = semanticTab && wordTabPaintBox(semanticTab);
            const style = getComputedStyle(text.parentElement!);
            const size = parseFloat(style.fontSize);
            if (tab && tab.getBoundingClientRect().height === 0 && tab.dataset.wordTabMeasured === 'true') {
              if (tab.hasAttribute('data-word-tab-decoration')) {
                // A print fragment can have a different origin/zoom from the
                // editor. Preserve its own fractional tab paint coordinates.
                const rectangle = tab.getBoundingClientRect();
                const left = (rectangle.left - bounds.left) / scale;
                tab.style.setProperty('--word-tab-decoration-fraction-x', `${left - Math.floor(left)}px`);
                tab.style.setProperty('--word-tab-decoration-width', String(rectangle.width / scale));
              }
              // A measured zero-height tab box is the baseline itself. Its
              // whitespace Range has no painted glyph and cannot supply a
              // font ascent; rebasing that Range displaces decoration ink.
              shift += baseline - sourceTop - (tab.getBoundingClientRect().top - bounds.top) / scale;
              const leaderSize = parseFloat(getComputedStyle(tab).getPropertyValue('--word-tab-leader-size'));
              if (tab.dataset.wordTabLeaderPainted === 'true' && leaderSize > 0)
                tab.style.setProperty('--word-tab-leader-baseline',
                  `${wordBrowserFontMetrics(getComputedStyle(tab), leaderSize, 0, scale, cache).ascent}px`);
            } else if (size > 0) {
              const metrics = wordBrowserFontMetrics(
                style,
                size,
                parseFloat(style.lineHeight) || size * 1.2,
                scale,
                cache,
              );
              const range = document.createRange();
              range.setStart(text, 0);
              range.setEnd(text, 1);
              const actual =
                (range.getBoundingClientRect().top - bounds.top) / scale + metrics.rangeAscent;
              shift += baseline - sourceTop - actual;
            }
          }
        }
        if (!span.style.transform) {
          span.style.top = `${shift}px`;
          span.style.setProperty('--word-inline-paint-shift', `${shift}px`);
          span.style.setProperty('--word-baseline-shift', `${shift}px`);
          continue;
        }
        const previousShift = parseFloat(span.style.getPropertyValue('--word-baseline-shift'));
        const top = (span.getBoundingClientRect().top - bounds.top) / scale - previousShift;
        const origin =
          Math.round((Math.ceil(top * scale - 1e-8) - top * scale) * 64) / (scale * 64);
        span.style.top = `${origin}px`;
        span.style.transform = `translateY(${shift - origin}px)`;
        span.style.setProperty('--word-paint-origin', `${origin}px`);
        span.style.setProperty('--word-inline-paint-shift', `${origin}px`);
        span.style.setProperty('--word-baseline-shift', `${shift}px`);
      }
    }
  } finally {
    root.style.removeProperty('display');
  }
}
