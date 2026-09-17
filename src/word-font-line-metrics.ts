import { wordHyphenText } from './word-hyphen';
import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { wordLineSpacing, type WordLineSpacing } from './word-line-spacing';
import { wordNativeBaseline } from './word-native-baseline';
import { wordTabPaintBox } from './word-tab-dom';
import { wordBrowserFontMetrics } from './word-browser-font-metrics';
import { wordInterMetrics } from './word-inter-metrics';
import { wordDefaults } from './word-defaults';
import { wordTabLayoutKey } from './word-tab-state';
import { wordScriptMetrics } from './word-script-metrics';
import { syncWordSelection } from './word-links';

type MeasuredRun = {
  from: number;
  to: number;
  height: number;
  size: number;
  space?: boolean;
  shift?: number;
  origin?: number;
  baseline?: number;
};

/** Cancel layout-grid fractions before choosing an integer paint origin.
 * Serialized CSS widths/transforms lose sub-grid precision; applying ceil to
 * their raw round-trip error can alternate forever between origins 0 and 1. */
export function wordInlinePaintOrigin(top: number, scale: number) {
  const physical = Math.round(top * scale * 64) / 64;
  return (Math.ceil(physical) - physical) / scale;
}

/** Native line advances are rounded to twips after applying the automatic multiple. */
export function wordUniformLineHeight(
  sizePx: number,
  normalRatio: number,
  spacing: WordLineSpacing,
) {
  const natural = sizePx * normalRatio;
  return spacing.rule === 'exact'
    ? spacing.line / 15
    : spacing.rule === 'atLeast'
      ? Math.max(spacing.line / 15, Math.round(natural * 15) / 15)
      : Math.round(natural * (spacing.line / 240) * 15) / 15;
}

/** One baseline, one font descriptor, varying sizes. Nonempty Word lines use
 * their text metrics; the terminal paragraph mark supplies empty-line metrics.
 * Automatic extra leading follows the line; minimum extra leading precedes it.
 */
export function wordMixedLineMetrics(
  sizes: number[],
  ratio: number,
  spacing: WordLineSpacing,
  heightFor = (size: number, rule: WordLineSpacing) => wordUniformLineHeight(size, ratio, rule),
) {
  if (
    !sizes.length ||
    sizes.some((s) => !Number.isFinite(s) || s <= 0) ||
    !Number.isFinite(ratio) ||
    ratio <= 0 ||
    ratio >= 4 ||
    spacing.rule === 'exact' ||
    (spacing.rule === 'auto' && spacing.line < 240)
  )
    return null;
  const heights = sizes.map((s) => heightFor(s, { rule: 'auto', line: 240 }));
  const natural = Math.max(...heights);
  const target = heightFor(Math.max(...sizes), spacing);
  const extra = Math.max(0, target - natural);
  return {
    heights,
    before: spacing.rule === 'atLeast' ? extra : 0,
    after: spacing.rule === 'auto' ? extra : 0,
  };
}

export const wordFontLineMetricsKey = new PluginKey<DecorationSet>('wordFontLineMetrics');
export const wordLayoutChangeEvent = 'wordlayoutchange';
/** Chromium quantizes each line independently to its layout grid. Quantize
 * cumulative boundaries instead, so long zoomed paragraphs retain their total
 * native advance rather than losing the same fraction on every hard line. */
export function wordLineDisplayHeight(top: number, height: number, scale: number) {
  if (!(scale > 0 && Number.isFinite(scale))) return height;
  const grid = scale * 64;
  // Stay on the upper side of the chosen grid point when CSS converts a
  // fractional zoom back to device layout units; the epsilon never reaches
  // another layout unit and is not part of the semantic advance.
  return (Math.floor((top + height) * grid + 1e-8) - Math.floor(top * grid + 1e-8)) / grid + 1e-6;
}
const key = wordFontLineMetricsKey;
export function wordFontSizePixels(value: unknown): number | null {
  if (typeof value !== 'string') return null;
  const match = /^(\d+(?:\.\d+)?)(pt|px)$/.exec(value);
  if (!match) return null;
  const size = Number(match[1]) * (match[2] === 'pt' ? 4 / 3 : 1);
  return Number.isFinite(size) && size > 0 ? size : null;
}

/** Rectangles from one text node share font metrics. Separate physical rows
 * can still overlap vertically when glyph bounds exceed the line advance. */
export function wordTextRectsShareLine(rects: { top: number; width: number }[]) {
  const painted = rects.filter((rect) => rect.width > 0);
  return (
    painted.length > 0 &&
    Math.max(...painted.map((rect) => rect.top)) - Math.min(...painted.map((rect) => rect.top)) <= 1
  );
}
/** View-only leading for homogeneous Latin text, empty paragraph marks and
 * bounded mixed sizes. Mixed soft lines use measured baseline struts; objects,
 * mixed descriptors, empty mixed logical lines and complex scripts retain their fallback.
 * Semantic spacing values remain in the document; no measured pixels are serialized.
 */
export const WordFontLineMetrics = Extension.create<{ authoredDefaults: () => boolean }>({
  name: 'wordFontLineMetrics',
  addOptions() {
    return { authoredDefaults: () => false };
  },
  addProseMirrorPlugins() {
    const options = this.options;
    const editor = this.editor;
    return [
      new Plugin({
        key,
        state: {
          init: () => DecorationSet.empty,
          // Old zero-width line widgets can become keyboard stops inside a
          // newly joined text run. Remove stale measurements in the editing
          // transaction; the view recomputes them for the new paragraph ranges.
          apply: (tr, prior) => tr.getMeta(key) ?? (tr.docChanged ? DecorationSet.empty : prior),
        },
        props: { decorations: (state) => key.getState(state) },
        view(view) {
          let frame = 0,
            destroyed = false,
            signature = '',
            settlingPass = 0;
          const cache = new Map<string, number>();
          const baselineCache = new Map<string, { ascent: number; rangeAscent: number }>();
          const descriptor = (s: CSSStyleDeclaration) =>
            [s.fontFamily, s.fontWeight, s.fontStyle, s.fontStretch].join('|');
          const normalRatio = (style: CSSStyleDeclaration) => {
            const font = descriptor(style);
            if (cache.has(font)) return cache.get(font)!;
            const probe = document.createElement('span');
            probe.style.cssText =
              'all:initial;position:fixed;left:-100000px;top:0;visibility:hidden;pointer-events:none;white-space:pre;font-size:16384px;line-height:normal;display:inline-block';
            probe.style.fontFamily = style.fontFamily;
            probe.style.fontWeight = style.fontWeight;
            probe.style.fontStyle = style.fontStyle;
            probe.style.fontStretch = style.fontStretch;
            probe.textContent = 'Hg';
            document.body.append(probe);
            // Chromium caps very large CSS font sizes. Divide by the effective size,
            // never the requested probe size, so that cap cannot shrink line advances.
            const ratio =
              probe.getBoundingClientRect().height / parseFloat(getComputedStyle(probe).fontSize);
            probe.remove();
            if (cache.size > 256) cache.clear();
            cache.set(font, ratio);
            return ratio;
          };
          const measure = () => {
            frame = 0;
            if (destroyed || view.composing) return;
            settlingPass++;
            const listMarkers = new Map<number, { baseline: number; height: number; shift: number }>();
            const specs: {
              from: number;
              to: number;
              height: number;
              scriptSingle?: boolean;
              paint?: {
                from: number;
                to: number;
                shift: number;
                transform?: boolean;
                origin?: number;
                baseline?: number;
                script?: NonNullable<ReturnType<typeof wordScriptMetrics>>;
              }[];
              uniform?: { natural: number; align: 'top' | 'bottom' };
              mixed?: {
                before: number;
                after: number;
                runs: MeasuredRun[];
              };
              hardLines?: {
                from: number;
                to: number;
                size: number;
                natural: number;
                height: number;
                displayHeight?: number;
                fitHeight?: number;
                align: 'top' | 'bottom';
                shift?: number;
                baseline?: number;
              }[];
              wrapped?: {
                struts: {
                  from: number;
                  height: number;
                  fitHeight: number;
                  depth: number;
                  lineFrom: number;
                  lineTo: number;
                }[];
                runs: MeasuredRun[];
              };
            }[] = [];
            let count = 0;
            let characters = 0;
            // Measurements consume rectangles immediately; no live Range enters
            // a layout spec. Reuse one instead of retaining one per text run.
            const range = document.createRange();
            view.state.doc.descendants((node, pos, parent) => {
              if (!['paragraph', 'heading'].includes(node.type.name)) return;
              if (++count > 5000 || /[^\u0000-\u024f]/.test(node.textContent)) return false;
              // Older authored paragraphs can inherit Normal entirely through
              // CSS. Resolve those same semantic defaults for measurements,
              // leaving the saved model and history untouched. Imported styles,
              // headings and nested formatting need their own inheritance.
              const inherit =
                options.authoredDefaults() &&
                parent === view.state.doc &&
                node.type.name === 'paragraph' &&
                !node.attrs.sourceParagraph;
              const paragraph = inherit
                ? {
                    ...node.attrs,
                    paragraphFontSize: node.attrs.paragraphFontSize ?? `${wordDefaults.fontSize}pt`,
                    paragraphFontFamily: node.attrs.paragraphFontFamily ?? wordDefaults.fontFamily,
                    paragraphLineHeight:
                      node.attrs.paragraphLineHeight ?? String(wordDefaults.lineMultiple),
                  }
                : node.attrs;
              const spacing = wordLineSpacing(
                paragraph.paragraphLineHeight,
                paragraph.paragraphLineRule,
              );
              if (!spacing) return false;
              const dom = view.nodeDOM(pos);
              if (!(dom instanceof HTMLElement)) return false;
              // A rejected or not-yet-measured tab cannot supply a reliable
              // physical line origin. Baseline paint otherwise moves these
              // fields repeatedly, causing tab and line measurement to reset
              // each other even after the settling-pass budget is exhausted.
              // Keep the semantic line spacing and ordinary editable fallback;
              // retry when tab layout publishes supported geometry.
              if (
                [...dom.querySelectorAll<HTMLElement>('[data-word-tab]')].some(
                  (tab) => wordTabPaintBox(tab).dataset.wordTabMeasured !== 'true',
                )
              )
                return false;
              let simple = true;
              let allText = true;
              let textOrTabs = true;
              node.forEach((child) => {
                if (
                  !child.isText &&
                  ![
                    'hardBreak',
                    'wordPageBreak',
                    'wordColumnBreak',
                    'wordTab',
                    'wordHyphen',
                  ].includes(child.type.name)
                )
                  simple = false;
                if (!child.isText && child.type.name !== 'wordHyphen') allText = false;
                if (!child.isText && !['wordTab', 'wordHyphen'].includes(child.type.name))
                  textOrTabs = false;
              });
              if (!simple) return false;
              const logicalText = node.textBetween(0, node.content.size, '', (leaf) =>
                leaf.type.name === 'wordTab'
                  ? '\t'
                  : leaf.type.name === 'wordHyphen'
                    ? wordHyphenText(leaf.attrs.kind) || ''
                    : '\n',
              );
              const nonemptyHardLines =
                !allText && logicalText.split('\n').every((line) => /\S/.test(line));
              let paragraphSize = paragraph.paragraphFontSize;
              let style = getComputedStyle(dom);
              if (!paragraphSize || (spacing.rule === 'exact' && !paragraph.paragraphFontFamily)) {
                // A nonempty exact line does not need an inferred paragraph-mark
                // font when all its text has the same explicit font and size.
                // Empty lines and unformatted/default fonts remain unsupported.
                if (
                  spacing.rule !== 'exact' ||
                  !logicalText.split('\n').every((line) => /\S/.test(line))
                )
                  return false;
                let explicit: { fontSize: string; fontFamily: string } | undefined;
                let known = true;
                node.forEach((child) => {
                  if (!child.isText) return;
                  const mark = child.marks.find((m) => m.type.name === 'textStyle')?.attrs;
                  if (
                    !mark?.fontSize ||
                    !mark.fontFamily ||
                    (explicit &&
                      (explicit.fontSize !== mark.fontSize ||
                        explicit.fontFamily !== mark.fontFamily))
                  )
                    known = false;
                  else explicit = { fontSize: mark.fontSize, fontFamily: mark.fontFamily };
                });
                if (!known || !explicit) return false;
                paragraphSize = explicit.fontSize;
                const text = document.createTreeWalker(dom, NodeFilter.SHOW_TEXT);
                while (text.nextNode()) {
                  if (
                    !text.currentNode.textContent ||
                    text.currentNode.parentElement?.closest('.ProseMirror-widget')
                  )
                    continue;
                  style = getComputedStyle(text.currentNode.parentElement!);
                  break;
                }
              }
              if (spacing.rule === 'exact' && [480, 800].includes(spacing.line)) {
                // The independently measured exact24/40 paragraphs include runs
                // whose uniform face differs from their paragraph-mark face.
                // Use that actual face only when all visible runs agree.
                const text = document.createTreeWalker(dom, NodeFilter.SHOW_TEXT);
                let first: Text | undefined, candidate: CSSStyleDeclaration | undefined;
                let uniform = true;
                while (text.nextNode()) {
                  if (
                    !text.currentNode.textContent ||
                    text.currentNode.parentElement?.closest('.ProseMirror-widget')
                  )
                    continue;
                  const current = getComputedStyle(text.currentNode.parentElement!);
                  if (!candidate) {
                    candidate = current;
                    first = text.currentNode as Text;
                  } else if (descriptor(current) !== descriptor(candidate)) uniform = false;
                }
                if (uniform && candidate && first) {
                  style = candidate;
                  const firstNode = view.state.doc.nodeAt(view.posAtDOM(first, 0));
                  paragraphSize =
                    firstNode?.marks.find((mark) => mark.type.name === 'textStyle')?.attrs
                      .fontSize || paragraphSize;
                }
              }
              const size = wordFontSizePixels(paragraphSize),
                font = descriptor(style);
              if (size === null) return false;
              const runs: {
                from: number;
                to: number;
                size: number;
                text: Text;
                script?: NonNullable<ReturnType<typeof wordScriptMetrics>>;
              }[] = [];
              let sameDescriptor = true,
                allBaselineAligned = true;
              const walker = document.createTreeWalker(dom, NodeFilter.SHOW_TEXT);
              while (walker.nextNode()) {
                if (walker.currentNode.parentElement?.closest('.ProseMirror-widget')) continue;
                if (!walker.currentNode.textContent) continue;
                const run = getComputedStyle(walker.currentNode.parentElement!);
                const from = view.posAtDOM(walker.currentNode, 0);
                const textNode = view.state.doc.nodeAt(from);
                const explicitSize = textNode?.marks.find((m) => m.type.name === 'textStyle')?.attrs
                  .fontSize;
                const runSize = wordFontSizePixels(explicitSize || paragraphSize);
                if (runSize === null) return false;
                const scriptMarks =
                  textNode?.marks.filter((m) =>
                    ['superscript', 'subscript'].includes(m.type.name),
                  ) || [];
                const script =
                  scriptMarks.length === 1 &&
                  (textNode?.isText ||
                    (textNode?.type.name === 'wordTab' && Math.abs(runSize * 0.75 - 10) < 0.0001))
                    ? wordScriptMetrics(
                        scriptMarks[0].type.name as 'superscript' | 'subscript',
                        runSize,
                        size,
                        run,
                        spacing,
                        allText,
                      )
                    : null;
                // Unknown scripts retain their editable fallback and explicit
                // PDF rejection, even if an imported CSS rule aligns them.
                if (scriptMarks.length && !script) return false;
                let baselineAligned = true;
                for (
                  let ancestor = walker.currentNode.parentElement;
                  ancestor && ancestor !== dom;
                  ancestor = ancestor.parentElement
                ) {
                  if (
                    !ancestor.hasAttribute('data-word-uniform-run-leading') &&
                    !ancestor.hasAttribute('data-word-hardline-run') &&
                    !ancestor.hasAttribute('data-word-space-leading') &&
                    !ancestor.classList.contains('word-reflow-line') &&
                    !(
                      script && ancestor.tagName === (script.kind === 'superscript' ? 'SUP' : 'SUB')
                    ) &&
                    getComputedStyle(ancestor).verticalAlign !== 'baseline'
                  ) {
                    baselineAligned = false;
                    allBaselineAligned = false;
                    sameDescriptor = false;
                    simple = false;
                  }
                }
                if (descriptor(run) !== font) sameDescriptor = false;
                runs.push({
                  from,
                  to: from + walker.currentNode.textContent!.length,
                  size: runSize,
                  text: walker.currentNode as Text,
                  ...(script ? { script } : {}),
                });
                if (runSize !== size || descriptor(run) !== font || !baselineAligned)
                  simple = false;
              }
              const ratio = normalRatio(style);
              if (!(ratio > 0 && ratio < 4)) return false;
              const candidate =
                simple && sameDescriptor ? wordInterMetrics(size, ratio, style) : null;
              const inter = candidate?.supports(spacing) ? candidate : null;
              const lineHeight = (runSize: number, rule: WordLineSpacing) =>
                inter && runSize === size
                  ? inter.height(rule)
                  : wordUniformLineHeight(runSize, ratio, rule);
              const nativeBaseline = (runSize: number, natural: number, rule: WordLineSpacing) =>
                inter && runSize === size
                  ? inter.ascent(rule)
                  : wordNativeBaseline(
                      runSize,
                      natural,
                      ratio,
                      style.fontFamily,
                      style.fontWeight,
                      style.fontStyle,
                      style.fontStretch,
                      rule,
                    );

              const canvas = dom.closest('.paper-wrap');
              const scale = canvas
                ? parseFloat(getComputedStyle(canvas).zoom)
                : dom.getBoundingClientRect().width / parseFloat(getComputedStyle(dom).width);
              const fontMetrics = (runSize: number) =>
                wordBrowserFontMetrics(
                  style,
                  runSize,
                  lineHeight(runSize, { rule: 'auto', line: 240 }),
                  scale,
                  baselineCache,
                );
              if (dom.hasAttribute('data-word-list-marker')) {
                const height = lineHeight(size, { rule: 'auto', line: 240 });
                const baseline = nativeBaseline(size, height, spacing);
                if (baseline !== null && scale > 0) listMarkers.set(pos, {
                  baseline, height, shift: baseline - fontMetrics(size).ascent,
                });
              }
              const paintOrigin = (text: Text) => {
                // Start transformed glyphs at an integer local paint origin,
                // then cancel that offset. This preserves fractional baselines
                // in Chromium PDF without moving the logical line or caret.
                const wrapper = text.parentElement?.closest(
                  '[data-word-run-leading],[data-word-baseline-paint]',
                );
                if (!wrapper) return 0;
                const prior = parseFloat(
                  getComputedStyle(wrapper).getPropertyValue('--word-baseline-shift'),
                );
                if (!Number.isFinite(prior)) return 0;
                const fragment = text.parentElement?.closest('.word-page-fragment');
                const displacement = fragment
                  ? parseFloat(getComputedStyle(fragment).getPropertyValue('--fragment-shift')) || 0
                  : 0;
                const top =
                  (wrapper.getBoundingClientRect().top - dom.getBoundingClientRect().top) / scale -
                  prior -
                  displacement;
                return wordInlinePaintOrigin(top, scale);
              };
              // Explicit, unwrapped lines can carry distinct empty-line fonts.
              // The break supplies an empty terminated line; the paragraph mark
              // supplies only an empty final line. Keep the source nodes/caret.
              if (
                sameDescriptor &&
                !textOrTabs &&
                node.childCount &&
                (spacing.rule !== 'auto' || spacing.line >= 240)
              ) {
                const lines: NonNullable<(typeof specs)[number]['hardLines']> = [];
                let from = pos + 1,
                  sizes: number[] = [],
                  valid = true,
                  hasBreak = false;
                const finish = (to: number, emptySize: number) => {
                  const lineSize = sizes[0] ?? emptySize;
                  if (sizes.some((s) => s !== lineSize)) valid = false;
                  const natural = lineHeight(lineSize, {
                    rule: 'auto',
                    line: 240,
                  });
                  const ascent = nativeBaseline(lineSize, natural, spacing);
                  const shift =
                    ascent !== null && scale > 0 && Number.isFinite(scale)
                      ? Math.round((ascent - fontMetrics(lineSize).ascent) * scale * 64) /
                        (scale * 64)
                      : undefined;
                  const height = lineHeight(lineSize, spacing);
                  if (
                    spacing.rule === 'exact' &&
                    (ascent === null || (lineSize > height && !inter))
                  )
                    valid = false;
                  lines.push({
                    from,
                    to,
                    size: lineSize,
                    natural,
                    height,
                    ...(spacing.rule === 'auto' && natural < height ? { fitHeight: natural } : {}),
                    align: spacing.rule === 'atLeast' ? 'bottom' : 'top',
                    shift,
                    ...(ascent === null
                      ? {}
                      : {
                          baseline:
                            lines.reduce((sum, line) => sum + line.height, 0) +
                            (spacing.rule === 'atLeast' ? Math.max(0, height - natural) : 0) +
                            ascent,
                        }),
                  });
                  from = to;
                  sizes = [];
                };
                node.forEach((child, offset) => {
                  const childSize = wordFontSizePixels(
                    child.marks.find((m) => m.type.name === 'textStyle')?.attrs.fontSize ||
                      paragraph.paragraphFontSize,
                  );
                  if (childSize === null) {
                    valid = false;
                    return;
                  }
                  if (child.isText) sizes.push(childSize);
                  else if (
                    ['hardBreak', 'wordPageBreak', 'wordColumnBreak'].includes(child.type.name)
                  ) {
                    const breakDOM = view.nodeDOM(pos + 1 + offset);
                    if (
                      !(breakDOM instanceof HTMLElement) ||
                      descriptor(getComputedStyle(breakDOM)) !== font ||
                      child.marks.some((m) => ['subscript', 'superscript'].includes(m.type.name))
                    )
                      valid = false;
                    hasBreak = true;
                    sizes.push(childSize);
                    finish(pos + 1 + offset + child.nodeSize, childSize);
                  } else valid = false;
                });
                // A standalone page-break paragraph before another paragraph
                // keeps its mark on the break line. Its successor starts the
                // new page; there is no additional empty line on that page.
                if (!(
                  node.childCount === 1 &&
                  node.firstChild?.type.name === 'wordPageBreak' &&
                  pos + node.nodeSize < view.state.doc.content.size
                ))
                  finish(pos + node.nodeSize - 1, size);
                // Wrapped or mixed-size logical lines still need the general
                // fragmentation engine. Never apply a one-line strut to them.
                let runIndex = 0;
                for (const line of lines) {
                  let top = -Infinity,
                    bottom = Infinity;
                  while (runIndex < runs.length && runs[runIndex].from < line.to) {
                    const run = runs[runIndex++];
                    if (run.from < line.from || run.to > line.to) valid = false;
                    range.selectNodeContents(run.text);
                    const rects = [...range.getClientRects()].filter((r) => r.width > 0);
                    if (rects.length !== 1) valid = false;
                    else {
                      top = Math.max(top, rects[0].top);
                      bottom = Math.min(bottom, rects[0].bottom);
                    }
                  }
                  if (top >= bottom) valid = false;
                }
                if (hasBreak && valid) {
                  // Computed widths are serialized with limited decimals.
                  // Their ratio can turn exact 50% into 0.500001 and floor
                  // every otherwise integral line one grid unit too short.
                  let top = 0;
                  for (const line of lines) {
                    line.displayHeight = wordLineDisplayHeight(top, line.height, scale);
                    top += line.height;
                  }
                  specs.push({
                    from: pos,
                    to: pos + node.nodeSize,
                    height: 0,
                    hardLines: lines,
                    paint: runs.flatMap((run) => {
                      const line = lines.find((l) => l.from <= run.from && l.to >= run.to);
                      return line?.shift === undefined
                        ? []
                        : [
                            {
                              from: run.from,
                              to: run.to,
                              shift: line.shift,
                              baseline: line.baseline,
                              transform: true,
                              origin: paintOrigin(run.text),
                            },
                          ];
                    }),
                  });
                  return false;
                }
              }
              const scriptSingle = spacing.rule === 'auto' && runs.some((run) => run.script);
              if (scriptSingle) {
                // Thirty native Arial10 Single controls preserve nominal line
                // height, including wholly scripted text/marks and empty marks.
                // Qualify one physical text line; tabs, hard breaks, wrapping
                // and mixed nominal fonts retain their separate guards.
                if (
                  !allText ||
                  !sameDescriptor ||
                  node.textContent.length > 4096 ||
                  runs.some((run) => run.size !== size)
                )
                  return false;
                let top = -Infinity,
                  bottom = Infinity;
                for (const run of runs) {
                  range.selectNodeContents(run.text);
                  const rects = [...range.getClientRects()].filter((rect) => rect.width > 0);
                  if (!wordTextRectsShareLine(rects)) return false;
                  for (const rect of rects) {
                    top = Math.max(top, rect.top);
                    bottom = Math.min(bottom, rect.bottom);
                  }
                  if (top >= bottom) return false;
                }
                if (!(top < bottom)) return false;
              }
              if (spacing.rule === 'exact' || scriptSingle) {
                const height = scriptSingle ? lineHeight(size, spacing) : spacing.line / 15;
                const ascent = nativeBaseline(size, size * ratio, spacing);
                // Mixed faces may share a measured exact-line baseline. Check
                // every actual descriptor instead of borrowing paragraph font
                // metrics for its descendants. Keep other line rules guarded.
                const runMetrics = new Map<Text, ReturnType<typeof wordBrowserFontMetrics>>();
                const compatibleMixed =
                  !sameDescriptor &&
                  allBaselineAligned &&
                  [480, 800].includes(spacing.line) &&
                  ascent !== null &&
                  runs.every((run) => {
                    const runStyle = getComputedStyle(run.text.parentElement!);
                    // Exact24 boundary-deletion references mix regular, bold
                    // and italic runs at one nominal font/size. Exact40 mixed
                    // family/size references retain their existing face guard.
                    if (spacing.line === 480) {
                      const family = (value: string) =>
                        value.split(',')[0].replace(/["']/g, '').trim().toLowerCase();
                      if (
                        run.size !== size ||
                        family(runStyle.fontFamily) !== family(style.fontFamily)
                      )
                        return false;
                    } else if (
                      runStyle.fontWeight !== style.fontWeight ||
                      runStyle.fontStyle !== style.fontStyle
                    )
                      return false;
                    const runRatio = normalRatio(runStyle);
                    const runAscent = wordNativeBaseline(
                      run.size,
                      run.size * runRatio,
                      runRatio,
                      runStyle.fontFamily,
                      runStyle.fontWeight,
                      runStyle.fontStyle,
                      runStyle.fontStretch,
                      spacing,
                    );
                    if (runAscent !== ascent) return false;
                    runMetrics.set(
                      run.text,
                      wordBrowserFontMetrics(
                        runStyle,
                        run.size,
                        run.size * runRatio,
                        scale,
                        baselineCache,
                      ),
                    );
                    return true;
                  });
                let oversizedArial24 = false;
                if (
                  spacing.rule === 'exact' &&
                  spacing.line === 480 &&
                  Math.abs(size * 0.75 - 40) < 0.0001 &&
                  style.fontFamily.split(',')[0].replace(/["']/g, '').trim().toLowerCase() ===
                    'arial' &&
                  sameDescriptor &&
                  allText &&
                  allBaselineAligned &&
                  ascent !== null &&
                  runs.length &&
                  runs.every((run) => run.size === size)
                ) {
                  let top = -Infinity,
                    bottom = Infinity;
                  for (const run of runs) {
                    range.selectNodeContents(run.text);
                    const rects = [...range.getClientRects()].filter((rect) => rect.width > 0);
                    if (!wordTextRectsShareLine(rects)) {
                      top = Infinity;
                      break;
                    }
                    for (const rect of rects) {
                      top = Math.max(top, rect.top);
                      bottom = Math.min(bottom, rect.bottom);
                    }
                  }
                  oversizedArial24 = top < bottom;
                }
                // Oversized glyph lines remain guarded outside the measured
                // Inter profile and single-line Arial40/exact24 control.
                if (
                  (!sameDescriptor && !compatibleMixed) ||
                  ascent === null ||
                  // Measured oversized cases keep their glyphs even when the
                  // font exceeds the line box. Other combinations stay gated.
                  ((size > height || runs.some((r) => r.size > height)) &&
                    !inter &&
                    !oversizedArial24) ||
                  !runs.length ||
                  !(scale > 0 && Number.isFinite(scale)) ||
                  settlingPass >= 8
                )
                  return false;
                const paint: NonNullable<(typeof specs)[number]['paint']> = [];
                if (
                  !dom.matches(
                    '[data-word-hardlines],[data-word-mixed-leading],[data-word-wrapped-leading]',
                  )
                )
                  for (const run of runs) {
                    const metrics = run.script
                      ? fontMetrics(parseFloat(getComputedStyle(run.text.parentElement!).fontSize))
                      : runMetrics.get(run.text) || fontMetrics(run.size);
                    const prior = run.text.parentElement?.closest('[data-word-baseline-paint]');
                    const oldShift = prior
                      ? parseFloat(
                          getComputedStyle(prior).getPropertyValue('--word-baseline-shift'),
                        ) || 0
                      : 0;
                    const fragment = run.text.parentElement?.closest('.word-page-fragment');
                    const displacement = fragment
                      ? parseFloat(
                          getComputedStyle(fragment).getPropertyValue('--fragment-shift'),
                        ) || 0
                      : 0;
                    range.setStart(run.text, 0);
                    range.setEnd(run.text, 1);
                    const semanticTab =
                      run.text.parentElement?.closest<HTMLElement>('[data-word-tab]');
                    const tabBounds =
                      semanticTab && wordTabPaintBox(semanticTab).getBoundingClientRect();
                    // A zero-leading tab is a baseline-aligned empty inline
                    // box. Its invisible character has no glyph ascent; at a
                    // trailing atom, Range can even expose an empty rectangle.
                    // The box's bottom supplies its actual inline baseline.
                    const baseline =
                      (tabBounds && tabBounds.height === 0
                        ? (tabBounds.bottom - dom.getBoundingClientRect().top) / scale
                        : (range.getBoundingClientRect().top - dom.getBoundingClientRect().top) /
                            scale +
                          metrics.rangeAscent) -
                      oldShift -
                      displacement;
                    // Unequal columns already have semantic line offsets from
                    // reflow. An overflowing inline box plus BR can introduce a
                    // browser-only row; it must not become the native baseline.
                    const reflow =
                      run.text.parentElement?.closest<HTMLElement>('[data-word-reflow-top]');
                    const reflowTop = reflow ? Number(reflow.dataset.wordReflowTop) : NaN;
                    // A measured oversized single line can have its browser
                    // baseline below the exact line box. That is not a second
                    // semantic line; move its ink without expanding the box.
                    const lineTop = oversizedArial24
                      ? 0
                      : Number.isFinite(reflowTop)
                        ? reflowTop
                        : Math.floor(baseline / height) * height;
                    const targetBaseline = lineTop + ascent + (run.script?.shift || 0);
                    const shift = targetBaseline - baseline;
                    range.setEnd(run.text, run.text.length);
                    const runBounds = range.getBoundingClientRect();
                    const fixedHardRun =
                      !!semanticTab ||
                      ((nonemptyHardLines ||
                        spacing.line === 800 ||
                        scriptSingle ||
                        oversizedArial24) &&
                        runBounds.height / scale <= metrics.rangeAscent + run.size &&
                        runBounds.width / scale <= parseFloat(getComputedStyle(dom).width) &&
                        new Set(
                          [...range.getClientRects()].map((rect) => Math.round(rect.top * 64)),
                        ).size === 1);
                    paint.push({
                      from: run.from,
                      to: run.to,
                      shift: Math.round(shift * scale * 64) / (scale * 64),
                      baseline: targetBaseline,
                      ...(run.script ? { script: run.script } : {}),
                      ...(fixedHardRun ? { transform: true, origin: paintOrigin(run.text) } : {}),
                    });
                  }
                specs.push({ from: pos, to: pos + node.nodeSize, height, paint, scriptSingle });
                return false;
              }
              // Single-line uniform text needs the same excess-leading placement
              // as mixed sizes. CSS otherwise centers extra spacing around it.
              let singleLine =
                sameDescriptor &&
                runs.length > 0 &&
                textOrTabs &&
                runs.every((r) => r.to - r.from <= 8192);
              if (singleLine) {
                let top = -Infinity,
                  bottom = Infinity;
                for (const run of runs) {
                  if (run.text.parentElement?.closest('[data-word-tab]')) continue;
                  range.selectNodeContents(run.text);
                  const rects = [...range.getClientRects()].filter((r) => r.width > 0);
                  if (!wordTextRectsShareLine(rects)) {
                    singleLine = false;
                    break;
                  }
                  // Chromium can split a leading space from its adjacent text
                  // into separate rectangles on the same line. Count physical
                  // rows, not rectangles, or our own wrappers alternate between
                  // uniform and mixed leading and printing never settles.
                  for (const rect of rects) {
                    top = Math.max(top, rect.top);
                    bottom = Math.min(bottom, rect.bottom);
                  }
                  // Intersected line boxes cannot overlap again after this
                  // becomes empty. Avoid scanning every remaining run/range in
                  // a long paragraph once wrapping has already been proved.
                  if (top >= bottom) {
                    singleLine = false;
                    break;
                  }
                }
                singleLine &&= top < bottom;
              }
              const metrics = singleLine
                ? wordMixedLineMetrics(
                    runs.map((r) => r.size),
                    ratio,
                    spacing,
                    lineHeight,
                  )
                : null;
              if (metrics) {
                if (!(scale > 0 && Number.isFinite(scale))) return false;
                const maximum = Math.max(...runs.map((r) => r.size));
                const metric = fontMetrics(maximum);
                // An inline-block paint wrapper changes mixed-size descendant
                // struts. Keep their proven inline advance path until that
                // separate baseline/paragraph-mark combination is measured.
                const ascent = runs.every((r) => r.size === size)
                  ? nativeBaseline(maximum, metric.natural, spacing)
                  : null;
                const shift =
                  ascent === null
                    ? undefined
                    : Math.round((ascent - metric.ascent) * scale * 64) / (scale * 64);
                specs.push({
                  from: pos,
                  to: pos + node.nodeSize,
                  height: 0,
                  mixed: {
                    before: metrics.before,
                    after: metrics.after,
                    runs: runs.map((r, i) => ({
                      from: r.from,
                      to: r.to,
                      height: metrics.heights[i],
                      size: r.size,
                      ...(shift === undefined
                        ? {}
                        : {
                            shift,
                            origin: paintOrigin(r.text),
                            baseline: metrics.before + ascent!,
                          }),
                    })),
                  },
                });
                return false;
              }
              // Soft and nonempty hard lines retain original text nodes and breaks.
              // A break's font does not contribute to a nonempty native line;
              // its source marks remain intact for export and empty-line editing.
              // Baseline-aligned zero-width struts supply the leading of each
              // measured line; they never become semantic breaks or stored data.
              if (
                sameDescriptor &&
                (allText || nonemptyHardLines) &&
                (!simple ||
                  nonemptyHardLines ||
                  nativeBaseline(size, lineHeight(size, { rule: 'auto', line: 240 }), spacing) !==
                    null) &&
                settlingPass < 8 &&
                runs.length &&
                node.textContent.length <= 4096 &&
                characters + node.textContent.length <= 10000 &&
                (spacing.rule !== 'auto' || spacing.line >= 240)
              ) {
                characters += node.textContent.length;
                const scale = dom.getBoundingClientRect().width / parseFloat(style.width);
                if (!(scale > 0 && Number.isFinite(scale))) return false;
                const lines: { from: number; baseline: number; sizes: number[] }[] = [];
                for (const run of runs) {
                  const metric = fontMetrics(run.size);
                  const fragment = run.text.parentElement?.closest('.word-page-fragment');
                  const fragmentShift = fragment
                    ? parseFloat(getComputedStyle(fragment).getPropertyValue('--fragment-shift')) ||
                      0
                    : 0;
                  const priorShift =
                    parseFloat(
                      getComputedStyle(run.text.parentElement!).getPropertyValue(
                        '--word-baseline-shift',
                      ),
                    ) || 0;
                  for (let i = 0; i < run.text.length; i++) {
                    if (/\s/.test(run.text.data[i])) continue;
                    range.setStart(run.text, i);
                    range.setEnd(run.text, i + 1);
                    const baseline =
                      range.getBoundingClientRect().top / scale +
                      metric.rangeAscent -
                      priorShift -
                      fragmentShift;
                    let line = lines.at(-1);
                    if (!line || Math.abs(line.baseline - baseline) > 1) {
                      line = { from: run.from + i, baseline, sizes: [] };
                      lines.push(line);
                    }
                    if (!line.sizes.includes(run.size)) line.sizes.push(run.size);
                  }
                }
                // A split inside a word needs shaping-aware fragments. Do not
                // introduce widget boundaries into that unsupported case.
                if (
                  lines.length > 1 &&
                  lines.every(
                    (line) =>
                      line.from === pos + 1 || /[ \t\n]/.test(logicalText[line.from - pos - 2]),
                  )
                ) {
                  const paragraphTop = dom.getBoundingClientRect().top / scale;
                  let lineTop = paragraphTop;
                  const baselines: (number | undefined)[] = [];
                  const shifts = lines.map((line) => {
                    const maximum = Math.max(...line.sizes);
                    const metric = fontMetrics(maximum);
                    const height = lineHeight(maximum, spacing);
                    const ascent = nativeBaseline(maximum, metric.natural, spacing);
                    const before =
                      spacing.rule === 'atLeast' ? Math.max(0, height - metric.natural) : 0;
                    const shift =
                      ascent === null ? undefined : lineTop + before + ascent - line.baseline;
                    baselines.push(
                      ascent === null ? undefined : lineTop - paragraphTop + before + ascent,
                    );
                    lineTop += height;
                    // Round in the zoomed layout grid so measurement converges
                    // at 50% as well as 100%, within the same bounded pass count.
                    return shift === undefined
                      ? undefined
                      : Math.round(shift * scale * 64) / (scale * 64);
                  });
                  const nativeBaselines = shifts.every((shift) => shift !== undefined);
                  specs.push({
                    from: pos,
                    to: pos + node.nodeSize,
                    height: 0,
                    wrapped: {
                      struts: lines.map((line, index) => {
                        const maximum = Math.max(...line.sizes);
                        const metric = fontMetrics(maximum);
                        const height = lineHeight(maximum, spacing);
                        const before =
                          spacing.rule === 'atLeast' ? Math.max(0, height - metric.natural) : 0;
                        // The first word owns the strut. A zero-width widget at
                        // the start could remain after the previous hanging space.
                        const word = /^\S+/.exec(logicalText.slice(line.from - pos - 1))![0];
                        return {
                          from: line.from + word.length,
                          height,
                          fitHeight:
                            spacing.rule === 'auto'
                              ? lineHeight(maximum, { rule: 'auto', line: 240 })
                              : height,
                          depth: height - metric.ascent - before,
                          lineFrom: index === 0 ? 0 : line.from - pos - 1,
                          lineTo: (lines[index + 1]?.from ?? pos + node.nodeSize - 1) - pos - 1,
                        };
                      }),
                      // Word excludes space-only font runs from this nonempty
                      // line's height, but their original advance width survives.
                      runs: runs.flatMap((r) =>
                        [...r.text.data.matchAll(/[ \t]+|[^ \t]+/g)].map((part) => ({
                          from: r.from + part.index!,
                          to: r.from + part.index! + part[0].length,
                          size: r.size,
                          // The strut owns the line box. Letting each inline run
                          // contribute another rounded box accumulates zoom error.
                          height: 0,
                          space: /^[ \t]+$/.test(part[0]),
                          ...(nativeBaselines
                            ? {
                                baseline:
                                  baselines[
                                    Math.max(
                                      0,
                                      lines.findLastIndex(
                                        (line) => line.from <= r.from + part.index!,
                                      ),
                                    )
                                  ],
                              }
                            : {}),
                          ...(nativeBaselines && !/^[ \t]+$/.test(part[0])
                            ? {
                                shift:
                                  shifts[
                                    lines.findLastIndex((line) => line.from <= r.from + part.index!)
                                  ],
                                origin: paintOrigin(r.text),
                              }
                            : {}),
                        })),
                      ),
                    },
                  });
                  return false;
                }
              }
              // Uniform fragments can align within each natural CSS line box;
              // no per-character positioning or synthetic editable lines needed.
              // Empty hard-break lines and mixed sizes still require more evidence.
              if (!simple) return false;
              const uniformText = node.textBetween(0, node.content.size, '', '\n');
              const natural = lineHeight(size, { rule: 'auto', line: 240 });
              specs.push({
                from: pos,
                to: pos + node.nodeSize,
                height: lineHeight(size, spacing),
                ...(runs.length &&
                uniformText.trim() &&
                !/^\n|\n$|\n\n/.test(uniformText) &&
                (spacing.rule !== 'auto' || spacing.line >= 240)
                  ? {
                      uniform: {
                        natural,
                        align: spacing.rule === 'atLeast' ? 'bottom' : 'top',
                      } as const,
                    }
                  : {}),
              });
              return false;
            });
            const next = JSON.stringify([specs, [...listMarkers]]);
            if (next === signature) return;
            signature = next;
            // An arrow can move the DOM caret before selectionchange reaches
            // ProseMirror. Replacing baseline wrappers must retain that move.
            syncWordSelection(editor);
            view.dispatch(
              view.state.tr
                .setMeta(
                  key,
                  DecorationSet.create(
                    view.state.doc,
                    specs.flatMap((s) => [
                      Decoration.node(s.from, s.to, {
                        'data-word-measured-leading': 'true',
                        ...(listMarkers.has(s.from) ? { 'data-word-list-baseline': String(listMarkers.get(s.from)!.baseline) } : {}),
                        ...(s.paint?.some((r) => r.script)
                          ? { 'data-word-script-layout': 'true' }
                          : {}),
                        ...(s.scriptSingle ? { 'data-word-script-single-line': 'true' } : {}),
                        ...(s.mixed ? { 'data-word-mixed-leading': 'true' } : {}),
                        ...(s.hardLines ? { 'data-word-hardlines': 'true' } : {}),
                        ...(s.wrapped ? { 'data-word-wrapped-leading': 'true' } : {}),
                        style:
                          `--word-measured-leading:${s.height}px` +
                          (listMarkers.has(s.from) ? `;--word-list-marker-shift:${listMarkers.get(s.from)!.shift}px;--word-list-marker-height:${listMarkers.get(s.from)!.height}px` : '') +
                          (s.paint?.length ||
                          s.mixed?.runs.some((r) => r.shift !== undefined) ||
                          s.wrapped?.runs.some((r) => r.shift !== undefined)
                            ? ';transform:translateZ(0)'
                            : '') +
                          (s.mixed
                            ? `;--word-mixed-before:${s.mixed.before}px;--word-mixed-after:${s.mixed.after}px`
                            : ''),
                      }),
                      ...(!s.hardLines
                        ? (s.paint?.map((r) =>
                            Decoration.inline(r.from, r.to, {
                              'data-word-baseline-paint': 'true',
                              ...(r.script ? { 'data-word-script-paint': r.script.kind } : {}),
                              ...(r.baseline === undefined
                                ? {}
                                : { 'data-word-native-baseline': String(r.baseline) }),
                              style:
                                (r.script ? `font-size:${r.script.size}px!important;` : '') +
                                (r.transform
                                  ? `--word-baseline-shift:${r.shift}px;--word-inline-paint-shift:${r.origin || 0}px;display:inline-block;white-space:pre;position:relative;top:${r.origin || 0}px;transform:translateY(${r.shift - (r.origin || 0)}px)`
                                  : `--word-baseline-shift:${r.shift}px;--word-inline-paint-shift:${r.shift}px;position:relative;top:${r.shift}px`),
                            }),
                          ) ?? [])
                        : []),
                      ...((s.mixed?.runs ?? s.wrapped?.runs)?.map((r) =>
                        Decoration.inline(r.from, r.to, {
                          'data-word-run-leading': 'true',
                          ...('baseline' in r && r.baseline !== undefined
                            ? { 'data-word-native-baseline': String(r.baseline) }
                            : {}),
                          ...('space' in r && r.space ? { 'data-word-space-leading': 'true' } : {}),
                          style:
                            `--word-run-leading:${r.height}px;--word-run-size:${r.size}px` +
                            ('space' in r && r.space ? ';vertical-align:top!important' : '') +
                            ('shift' in r && r.shift !== undefined
                              ? `;--word-baseline-shift:${r.shift}px;--word-paint-origin:${r.origin || 0}px;display:inline-block;position:relative;top:${r.origin || 0}px;transform:translateY(${r.shift - (r.origin || 0)}px)`
                              : ''),
                        }),
                      ) ?? []),
                      ...(s.wrapped?.struts.map((line) =>
                        Decoration.widget(
                          line.from,
                          () => {
                            const strut = document.createElement('span');
                            strut.setAttribute('data-word-softline-strut', 'true');
                            strut.setAttribute('data-word-line-from', String(line.lineFrom));
                            strut.setAttribute('data-word-line-to', String(line.lineTo));
                            if (line.fitHeight < line.height)
                              strut.setAttribute(
                                'data-word-line-fit-height',
                                String(line.fitHeight),
                              );
                            strut.setAttribute('aria-hidden', 'true');
                            strut.style.cssText = `display:inline-block;width:0;height:${line.height}px;vertical-align:${-line.depth}px;pointer-events:none`;
                            return strut;
                          },
                          // Keep the strut outside the preceding transformed word.
                          { side: 1, ignoreSelection: true },
                        ),
                      ) ?? []),
                      ...(s.hardLines?.flatMap((line) => [
                        Decoration.widget(
                          line.from,
                          () => {
                            const strut = document.createElement('span');
                            strut.setAttribute('data-word-line-strut', 'true');
                            strut.setAttribute(
                              'data-word-line-from',
                              String(line.from - s.from - 1),
                            );
                            strut.setAttribute('data-word-line-to', String(line.to - s.from - 1));
                            strut.setAttribute('data-word-line-advance', String(line.height));
                            if (line.fitHeight !== undefined)
                              strut.setAttribute(
                                'data-word-line-fit-height',
                                String(line.fitHeight),
                              );
                            strut.setAttribute('aria-hidden', 'true');
                            strut.style.cssText = `display:inline-block;width:0;height:${line.displayHeight ?? line.height}px;vertical-align:top;pointer-events:none`;
                            return strut;
                          },
                          { side: -1, ignoreSelection: true },
                        ),
                        ...(() => {
                          // Combine line geometry and paint on the same range.
                          // Independent overlapping decorations can swap their
                          // nesting across a font/edit-session mark boundary,
                          // introducing a second strut and oscillating layout.
                          const paints = (s.paint || []).filter(
                            (r) => r.from >= line.from && r.to <= line.to,
                          );
                          const points = [
                            ...new Set([
                              line.from,
                              line.to,
                              ...paints.flatMap((r) => [r.from, r.to]),
                            ]),
                          ].sort((a, b) => a - b);
                          return points.slice(0, -1).map((from, i) => {
                            const to = points[i + 1];
                            const paint = paints.find((r) => r.from <= from && r.to >= to);
                            const origin = paint?.origin || 0;
                            return Decoration.inline(from, to, {
                              'data-word-hardline-run': 'true',
                              ...(paint ? { 'data-word-baseline-paint': 'true' } : {}),
                              ...(paint?.baseline === undefined
                                ? {}
                                : { 'data-word-native-baseline': String(paint.baseline) }),
                              style:
                                // Each accepted hard line is already unwrapped.
                                // Preserve even a whitespace-only paint piece:
                                // pre-wrap otherwise hangs it at zero width in
                                // the print clone, closing an authored word gap.
                                `font-size:${line.size}px;white-space:pre;line-height:${Math.min(line.natural, line.displayHeight ?? line.natural)}px!important;vertical-align:${line.align}!important` +
                                (paint
                                  ? `;--word-baseline-shift:${paint.shift}px;--word-inline-paint-shift:${origin}px;display:inline-block;position:relative;top:${origin}px;transform:translateY(${paint.shift - origin}px)`
                                  : ''),
                            });
                          });
                        })(),
                      ]) ?? []),
                      ...(s.uniform
                        ? [
                            Decoration.inline(s.from + 1, s.to - 1, {
                              'data-word-uniform-run-leading': 'true',
                              style: `--word-uniform-natural:${s.uniform.natural}px;--word-uniform-align:${s.uniform.align}`,
                            }),
                          ]
                        : []),
                    ]),
                  ),
                )
                .setMeta('addToHistory', false),
            );
            // Mapped decorations may still carry the previous font size or line
            // split during the first measure after an edit. Measure the resulting
            // DOM again even when the page's fixed minimum height hides its resize.
            schedule(false);
          };
          const schedule = (reset: unknown = true) => {
            if (reset !== false) settlingPass = 0;
            if (!destroyed && !frame) frame = requestAnimationFrame(measure);
          };
          const fontsChanged = () => {
            cache.clear();
            baselineCache.clear();
            signature = '';
            schedule();
          };
          document.fonts.addEventListener('loadingdone', fontsChanged);
          document.fonts.ready.then(schedule);
          let measuredWidth = -1;
          const resize = new ResizeObserver(() => {
            const width = view.dom.getBoundingClientRect().width;
            if (width !== measuredWidth) {
              measuredWidth = width;
              schedule();
            }
          });
          resize.observe(view.dom);
          view.dom.addEventListener('compositionend', schedule);
          view.dom.addEventListener(wordLayoutChangeEvent, schedule);
          schedule();
          return {
            update(current, previous) {
              if (current.state.doc !== previous.doc) {
                signature = '';
                schedule();
              } else if (
                wordTabLayoutKey.getState(current.state) !== wordTabLayoutKey.getState(previous)
              ) {
                schedule();
              }
            },
            destroy() {
              destroyed = true;
              cancelAnimationFrame(frame);
              resize.disconnect();
              document.fonts.removeEventListener('loadingdone', fontsChanged);
              view.dom.removeEventListener('compositionend', schedule);
              view.dom.removeEventListener(wordLayoutChangeEvent, schedule);
            },
          };
        },
      }),
    ];
  },
});
