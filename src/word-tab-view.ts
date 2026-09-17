import { Extension } from '@tiptap/core';
import { Plugin, TextSelection } from '@tiptap/pm/state';
import { DOMSerializer, type Node as WordNode } from '@tiptap/pm/model';
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view';
import { wordTabStops, wordDefaultTab, wordDecimalSymbol } from './word-tab-stops';
import { wordTabAdvance } from './word-tab-layout';
import { wordLayoutChangeEvent, wordFontSizePixels } from './word-font-line-metrics';
import { measureWordLines } from './word-line-measurements';
import { syncWordSelection } from './word-links';
import { wordTabBarOffsets } from './word-tab-bars';
import { wordTabLayoutKey, type WordTabLayoutState } from './word-tab-state';
import { wordTabLeaderOrigins, wordTabLeaderPrintBounds } from './word-tab-leader-layout';
import { wordTabLeaderFont } from './word-tab-leader-font';
import { wordTabLeaderNormalRatio } from './word-tab-leader-metrics';
import { wordTabPaintBox } from './word-tab-dom';
import { wordScriptMetrics } from './word-script-metrics';
import { wordLineSpacing } from './word-line-spacing';
import { wordTabDecoration } from './word-tab-decoration';
import { wordBrowserFontMetrics } from './word-browser-font-metrics';
export { wordTabLayoutKey } from './word-tab-state';

/** Chromium can stop Home/End at an atomic tab's paint wrapper. Use the
 * verified single-line semantic range; wrapped-line affinity remains native. */
function tabLineEdge(view: EditorView, end: boolean, extend: boolean) {
  if (view.composing || !(view.state.selection instanceof TextSelection)) return false;
  const selection = view.state.selection, head = selection.$head;
  if (!head.depth || !head.parent.isTextblock) return false;
  const from = head.before(), paragraph = view.nodeDOM(from);
  // Keyboard input can arrive before the asynchronous spacer measurement.
  // A verified single rendered line already establishes its semantic edges;
  // requiring a measured spacer lets native Home stop after a leading atom.
  if (!(paragraph instanceof HTMLElement) || !paragraph.querySelector('[data-word-tab]')) return false;
  const lines = measureWordLines(view, from);
  if (lines?.length !== 1) return false;
  const target = head.start() + (end ? head.parent.content.size : 0);
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, extend ? selection.anchor : target, target)).scrollIntoView());
  return true;
}


/** Prepare semantic fields without current baseline/fragment paint. All probes
 * are attached together, read together and removed before dispatching layout. */
const fieldFontProperties = ['font-family', 'font-size', 'font-weight', 'font-style', 'font-stretch', 'font-kerning', 'font-feature-settings', 'letter-spacing', 'word-spacing'];
type FieldStyle = (readonly [string, string])[];
type FieldCache = { style: string; widths: Map<string, number> };
function fieldProbe(view: EditorView, from: number, to: number, style: FieldStyle, serializer: DOMSerializer) {
  if (to <= from) return null;
  const box = document.createElement('span');
  box.style.cssText = 'position:fixed;left:-100000px;top:0;visibility:hidden;pointer-events:none;display:inline-block;white-space:pre;width:max-content';
  for (const [property, value] of style) box.style.setProperty(property, value);
  box.append(serializer.serializeFragment(view.state.doc.slice(from, to).content));
  return box;
}

export const WordTabView = Extension.create<{ pageLeft: () => number | null }>({
  name: 'wordTabView',
  addOptions: () => ({ pageLeft: () => null }),
  addKeyboardShortcuts() {
    return {
      Home: () => tabLineEdge(this.editor.view, false, false),
      End: () => tabLineEdge(this.editor.view, true, false),
      'Shift-Home': () => tabLineEdge(this.editor.view, false, true),
      'Shift-End': () => tabLineEdge(this.editor.view, true, true),
    };
  },
  addProseMirrorPlugins() {
    const editor = this.editor, pageLeft = this.options.pageLeft;
    return [new Plugin<WordTabLayoutState>({
      key: wordTabLayoutKey,
      state: {
        init: () => ({ signature: '', decorations: DecorationSet.empty }),
        apply: (tr, value) => tr.getMeta(wordTabLayoutKey) || (tr.docChanged
          ? { signature: '', decorations: value.decorations.map(tr.mapping, tr.doc) } : value),
      },
      props: { decorations: (state) => wordTabLayoutKey.getState(state)?.decorations },
      view(view) {
        let frame = 0, disposed = false, settlingPass = 0, exhausted = false, measuredWidth = -1;
        // Immutable paragraph identity includes every semantic field and mark.
        // Widths do not depend on the preceding tabs' current DOM positions.
        let fieldCache = new WeakMap<WordNode, FieldCache>();
        const leaderFonts = new Map<string, number>();
        const leaderBaselines = new Map<string, { ascent: number; rangeAscent: number }>();
        const measure = () => {
          frame = 0;
          if (disposed || exhausted || !view.dom.isConnected || view.composing) return;
          const specs: { from: number; to: number; attrs: Record<string, string> }[] = [];
          const probes: HTMLElement[] = [], finish: (() => void)[] = [];
          const widths = new Map<HTMLElement, number>();
          const pendingCache = new Map<HTMLElement, { cache: FieldCache; key: string; paragraph: WordNode }>();
          const serializer = DOMSerializer.fromSchema(view.state.schema);
          const probe = (from: number, to: number, position: number, style: FieldStyle, cache: FieldCache): HTMLElement | number => {
            if (to <= from) return 0;
            const key = `${from - position}:${to - position}`, cached = cache.widths.get(key);
            if (cached !== undefined) return cached;
            const element = fieldProbe(view, from, to, style, serializer)!;
            probes.push(element);pendingCache.set(element, { cache, key, paragraph: view.state.doc.nodeAt(position)! });
            return element;
          };
          let count = 0, leaderGlyphs = 0;
          const physicalLeft = pageLeft();
          view.state.doc.descendants((node, position) => {
            if (!node.isTextblock) return;
            const tabs: { from: number; to: number }[] = [];
            node.forEach((child, offset) => {
              if (child.type.name === 'wordTab') tabs.push({ from: position + 1 + offset, to: position + 2 + offset });
              else if (['hardBreak', 'wordPageBreak', 'wordColumnBreak'].includes(child.type.name))
                tabs.push({ from: position + 1 + offset, to: -1 });
            });
            const bars = wordTabBarOffsets(node.attrs.paragraphTabs);
            if (!tabs.some((tab) => tab.to > 0) && !bars.length) return false;
            const paragraph = view.nodeDOM(position);
            if (!(paragraph instanceof HTMLElement)) return false;
            const style = getComputedStyle(paragraph), bounds = paragraph.getBoundingClientRect();
            const canvas = paragraph.closest('.paper-wrap');
            const scale = canvas ? parseFloat(getComputedStyle(canvas).zoom) : bounds.width / parseFloat(style.width);
            if (!(scale > 0 && Number.isFinite(scale))) return false;
            const indent = parseFloat(style.marginInlineStart) || 0;
            if (bars.length && style.direction === 'ltr' && !node.attrs.paragraphTabUnsupported)
              specs.push({ from: position, to: position + node.nodeSize, attrs: {
                'data-word-bar-paint': 'true',
                'data-word-bar-offsets': JSON.stringify(bars.map((offset) => offset - indent)),
                // Chromium rounds a subpixel shadow box to a full paint pixel.
                // Paint a unit-width box and scale its complete coordinate
                // system afterward, keeping both ink width and stop positions.
                style: `--word-bar-shadows:${bars.map((offset) => `${(offset - indent) * 5}px 0 0 0 #000`).join(',')}`,
              } });
            const stops = wordTabStops(node.attrs.paragraphTabs) || [];
            const interval = wordDefaultTab(node.attrs.paragraphDefaultTab) ?? 720;
            const decimal = wordDecimalSymbol(node.attrs.paragraphDecimalSymbol) || '.';
            // Baseline paint can zero the computed size. Authored inline size
            // and semantic marks remain the typography used by field probes.
            const fieldStyle = fieldFontProperties.map((property) => [property,
              property === 'font-size' && paragraph.style.fontSize ? paragraph.style.fontSize : style.getPropertyValue(property)] as const);
            const styleKey = JSON.stringify(fieldStyle);
            let cached = fieldCache.get(node);
            if (!cached || cached.style !== styleKey) {
              cached = { style: styleKey, widths: new Map() };fieldCache.set(node, cached);
            }
            for (const [index, tab] of tabs.entries()) {
              if (tab.to < 0) continue;
              // Bound the expensive field serialization/layout itself. Checking
              // only after field measurement still processes every rejected tab.
              if (++count > 1000 || style.direction !== 'ltr' || node.attrs.paragraphTabUnsupported) {
                specs.push({ from: tab.from, to: tab.to, attrs: { 'data-word-tab-measured': 'false' } });
                continue;
              }
              const element = view.nodeDOM(tab.from);
              if (!(element instanceof HTMLElement)) continue;
              const dom = wordTabPaintBox(element);
              const end = tabs[index + 1]?.from ?? position + node.nodeSize - 1;
              const field = view.state.doc.textBetween(tab.to, end, '', '');
              const decimalAt = field.indexOf(decimal);
              const fragment = dom.closest('.word-page-fragment');
              const shift = fragment ? parseFloat(getComputedStyle(fragment).getPropertyValue('--fragment-left')) || 0 : 0;
              const current = (dom.getBoundingClientRect().left - bounds.left) / scale + indent - shift;
              const following = probe(tab.to, end, position, fieldStyle, cached);
              const prefix = decimalAt < 0 ? following : probe(tab.to, tab.to + decimalAt, position, fieldStyle, cached);
              finish.push(() => {
                const result = wordTabAdvance(current, typeof following === 'number' ? following : widths.get(following)!,
                  typeof prefix === 'number' ? prefix : widths.get(prefix)!, stops, interval, indent, indent + bounds.width / scale);
                if (result) {
                  const width = Math.max(0, Math.floor(result.width * scale * 64) / (scale * 64));
                  const attrs: Record<string, string> = {
                    'data-word-tab-measured': 'true',
                    'data-word-tab-leader': result.leader,
                    style: `display:inline-block;vertical-align:baseline;white-space:pre;tab-size:0;width:${width.toFixed(6)}px`,
                  };
                  const markedTab = view.state.doc.nodeAt(tab.from);
                  const marks = (markedTab?.marks.some(mark => mark.type.name === 'underline') ? 1 : 0)
                    | (markedTab?.marks.some(mark => mark.type.name === 'strike') ? 2 : 0);
                  // A qualified zero-height tab box is at its semantic
                  // baseline. Its whitespace has no text-decoration advance;
                  // paint the measured rectangle without adding a character.
                  if (marks && dom.closest('[data-word-native-baseline]') && dom.getBoundingClientRect().height === 0
                    && !markedTab?.marks.some(mark => ['superscript', 'subscript'].includes(mark.type.name))) {
                    const paintStyle = getComputedStyle(dom);
                    const size = wordFontSizePixels(markedTab?.marks.find(mark => mark.type.name === 'textStyle')?.attrs.fontSize
                      || node.attrs.paragraphFontSize || fieldStyle.find(([name]) => name === 'font-size')?.[1]);
                    const paint = size === null ? null : wordTabDecoration({ family: paintStyle.fontFamily, size: size * .75,
                      weight: paintStyle.fontWeight, style: paintStyle.fontStyle, stretch: paintStyle.fontStretch,
                      letterSpacing: paintStyle.letterSpacing, normalRatio: wordTabLeaderNormalRatio(paintStyle, leaderFonts) }, marks);
                    if (paint) {
                      const localLeft = current - indent + shift;
                      const fraction = localLeft - Math.floor(localLeft);
                      attrs['data-word-tab-decoration'] = String(marks);
                      attrs.style += `;--word-tab-decoration-top:${paint.top}px;--word-tab-decoration-height:${paint.paintHeight}px;`
                        + `--word-tab-decoration-background:${paint.background};--word-tab-decoration-fraction-x:${fraction}px;`
                        + `--word-tab-decoration-width:${width}`;
                    }
                  }
                  if (result.leader !== 'none' && physicalLeft !== null && Number.isFinite(physicalLeft)) {
                    const tabNode = view.state.doc.nodeAt(tab.from), tabStyle = getComputedStyle(dom);
                    const nominalSize = wordFontSizePixels(tabNode?.marks.find((mark) => mark.type.name === 'textStyle')?.attrs.fontSize
                      || node.attrs.paragraphFontSize || fieldStyle.find(([name]) => name === 'font-size')?.[1]);
                    const paragraphSize = wordFontSizePixels(node.attrs.paragraphFontSize);
                    const spacing = wordLineSpacing(node.attrs.paragraphLineHeight, node.attrs.paragraphLineRule);
                    const scripts = tabNode?.marks.filter(mark => ['superscript', 'subscript'].includes(mark.type.name)) || [];
                    const script = scripts.length === 1 && nominalSize !== null && paragraphSize !== null && spacing
                      ? wordScriptMetrics(scripts[0].type.name as 'superscript' | 'subscript', nominalSize, paragraphSize, tabStyle, spacing) : null;
                    const size = script?.size ?? nominalSize;
                    const font = size === null || (scripts.length > 0 && !script) ? null : wordTabLeaderFont(result.leader, {
                      family: tabStyle.fontFamily, size: size * .75, weight: tabStyle.fontWeight,
                      style: tabStyle.fontStyle, stretch: tabStyle.fontStretch, letterSpacing: tabStyle.letterSpacing,
                      normalRatio: wordTabLeaderNormalRatio(tabStyle, leaderFonts),
                      ...(script ? { script: script.kind } : {}),
                    });
                    const start = (physicalLeft + current) * .75;
                    const printBounds = wordTabLeaderPrintBounds(physicalLeft * .75, current * .75, result.width * .75);
                    const origins = font && printBounds ? wordTabLeaderOrigins(...printBounds, font.pitch) : null;
                    if (font && origins && leaderGlyphs + origins.length <= 5000) {
                      leaderGlyphs += origins.length;
                      const offsets = origins.map((x) => (x - start) / .75);
                      attrs['data-word-tab-leader-painted'] = 'true';
                      attrs['data-word-tab-leader-glyph'] = origins.length ? font.glyph : '';
                      attrs['data-word-tab-leader-offsets'] = JSON.stringify(offsets);
                      attrs.style += `;--word-tab-leader-size:${size}px;--word-tab-leader-first:${offsets[0] || 0}px;`
                        + `--word-tab-leader-shadows:${offsets.length > 1 ? offsets.slice(1).map((x) => `${x - offsets[0]}px 0 0 currentColor`).join(',') : 'none'}`;
                      if (size !== null && dom.closest('[data-word-native-baseline]') && dom.getBoundingClientRect().height === 0) {
                        const baseline = wordBrowserFontMetrics(tabStyle, size, 0, scale, leaderBaselines).ascent;
                        attrs.style += `;--word-tab-leader-baseline:${baseline}px`;
                      }
                    }
                  }
                  specs.push({ from: tab.from, to: tab.to, attrs });
                } else specs.push({ from: tab.from, to: tab.to, attrs: { 'data-word-tab-measured': 'false' } });
              });
            }
            return false;
          });
          // Avoid forcing whole-document layout once for every tab field.
          if (probes.length) {
            const holder = document.createElement('div');
            holder.dataset.wordTabProbes = 'true';
            holder.style.cssText = 'position:fixed;left:-100000px;top:0;visibility:hidden;pointer-events:none;contain:layout style size';
            for (const element of probes) holder.append(element);
            document.body.append(holder);
            try {
              const scripts: { element: HTMLElement; size: number }[] = [];
              for (const element of probes) {
                const paragraph = pendingCache.get(element)!.paragraph;
                const size = wordFontSizePixels(paragraph.attrs.paragraphFontSize);
                const spacing = wordLineSpacing(paragraph.attrs.paragraphLineHeight, paragraph.attrs.paragraphLineRule);
                if (size === null || !spacing) continue;
                for (const marker of element.querySelectorAll<HTMLElement>('sup,sub')) {
                  if (marker.querySelector('sup,sub') || marker.parentElement!.closest('sup,sub')) continue;
                  const style = getComputedStyle(marker.parentElement!);
                  const profile = wordScriptMetrics(marker.tagName === 'SUP' ? 'superscript' : 'subscript',
                    parseFloat(style.fontSize), size, style, spacing);
                  if (profile) scripts.push({ element: marker, size: profile.size });
                }
              }
              // Read the semantic probe fonts together, then apply the same
              // qualified paint sizes as the visible line before width reads.
              for (const script of scripts) {
                script.element.style.fontSize = script.size + 'px';
                script.element.style.verticalAlign = 'baseline';
              }
              for (const element of probes) {
                const width = element.getBoundingClientRect().width;
                widths.set(element, width);
                const pending = pendingCache.get(element)!;
                if (Number.isFinite(width)) pending.cache.widths.set(pending.key, width);
              }
            } finally {
              // One detach also avoids updating every live document Range once
              // for each field removed from the document body.
              holder.remove();
            }
          }
          for (const complete of finish) complete();
          let signature = JSON.stringify(specs);
          if (signature !== wordTabLayoutKey.getState(view.state)?.signature) {
            // Wrapped field positions can depend on preceding spacers from the
            // last paint. A document that does not converge must remain editable
            // without an endless chain of layout transactions or a false PDF
            // success. Retry after content, font, zoom or available width changes.
            if (++settlingPass > 32) {
              exhausted = true;
              for (const spec of specs) if ('data-word-tab-measured' in spec.attrs)
                spec.attrs = { 'data-word-tab-measured': 'false' };
              signature = JSON.stringify(specs);
            }
            // Native arrow navigation may have moved the DOM caret before its
            // selectionchange observer runs. Preserve that selection before a
            // spacer repaint can restore the older model caret.
            syncWordSelection(editor);
            view.dispatch(view.state.tr.setMeta(wordTabLayoutKey, {
              signature,
              decorations: DecorationSet.create(view.state.doc, specs.map((s) => Decoration.node(s.from, s.to, s.attrs))),
            }).setMeta('addToHistory', false));
          }
        };
        const schedule = () => { if (!frame && !disposed) frame = requestAnimationFrame(measure); };
        const reset = () => { settlingPass = 0;exhausted = false;schedule(); };
        const invalidate = () => { fieldCache = new WeakMap();leaderFonts.clear();leaderBaselines.clear();reset(); };
        const observer = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(() => {
          const width = view.dom.getBoundingClientRect().width;
          if (width !== measuredWidth) { measuredWidth = width;reset(); }
        });
        observer?.observe(view.dom);
        view.dom.addEventListener(wordLayoutChangeEvent, invalidate);
        view.dom.addEventListener('compositionend', schedule);
        document.fonts?.addEventListener('loadingdone', invalidate);
        schedule();
        return { update(current, previous) {
          if (current.state.doc !== previous.doc) reset();
          else schedule();
        }, destroy() {
          disposed = true; cancelAnimationFrame(frame); observer?.disconnect();
          view.dom.removeEventListener(wordLayoutChangeEvent, invalidate);
          view.dom.removeEventListener('compositionend', schedule);
          document.fonts?.removeEventListener('loadingdone', invalidate);
        } };
      },
    })];
  },
});
