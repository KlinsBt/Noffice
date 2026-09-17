import { Extension } from '@tiptap/core';
import { wordHyphenText } from './word-hyphen';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import type { WordContent } from './model';
import { wordSectionMap } from './word-editor-sections';
import { resolveWordSections } from './word-section-layout';
import { authoredWordSection, authoredWordSectionMap } from './word-authored-section';
import { measureWordPdf } from './word-pdf-measure';
import { wordStoryMeasurer } from './word-story-measure';
import type { WordPdfSnapshot } from './word-pdf-model';
import { measureWordLines, paragraphGraphemes } from './word-line-measurements';
import { wordFontLineMetricsKey, wordLayoutChangeEvent } from './word-font-line-metrics';
import { wordTabLayoutKey } from './word-tab-state';
import { wordTabPaintBox } from './word-tab-dom';
import { syncWordSelection } from './word-links';
import { wordSurfaceTabBars, wordTabBarOffsets } from './word-tab-bars';
import {
  planWordFragments,
  fragmentWordSurfaces,
  type WordFragmentPlan,
} from './word-fragment-plan';
import {
  reflowWordColumns,
  type WordColumnReflow,
  type WordLineProvider,
} from './word-column-reflow';
import { wordColumnLineProvider } from './word-column-line-provider';
import type { EditorView } from '@tiptap/pm/view';
import type { WordExportLayout } from './word-export-layout';
import { wordFragmentPrint, alignWordPrintGlyphs } from './word-fragment-print';
import { wordBodyGridOffsets } from './word-body-grid';
import {
  sectionSurfaceInput,
  planSectionSurfaces,
  sectionPrintRules,
  sectionPrintPage,
  type WordSurfacePlan,
  type SurfaceInput,
} from './word-section-surfaces';

export interface SurfaceViewState {
  plan: WordSurfacePlan | null;
  printRules: string;
  reason: string;
  storyHtml?: Record<string, string>;
}
const empty: SurfaceViewState = { plan: null, printRules: '', reason: '' };
export const wordSurfaceViewKey = new PluginKey<DecorationSet>('wordSurfaceView');
const key = wordSurfaceViewKey;
const pdfLayouts = new WeakMap<EditorView, (content: string) => WordPdfSnapshot>();
const pendingStories = new WeakSet<EditorView>();
export const wordStoryLayoutPending = (view: EditorView) => pendingStories.has(view);
export function wordPdfSnapshot(view: EditorView, content: string) {
  const capture = pdfLayouts.get(view);
  if (!capture) throw Error('The PDF layout is not ready.');
  return capture(content);
}
const exportLayouts = new WeakMap<
  EditorView,
  {
    doc: EditorView['state']['doc'];
    paragraphs: WordExportLayout['paragraphs'];
  }
>();
export function wordColumnExportParagraphs(view: EditorView) {
  const current = exportLayouts.get(view);
  return current?.doc === view.state.doc ? current.paragraphs : undefined;
}

/** View-only measurements/decorations: one contenteditable, no copied/moved text nodes,
 * no model transactions for layout, no geometry or print attributes in serialized HTML.
 */
export const WordSurfaceView = Extension.create<{
  content: () => WordContent;
  change: (state: SurfaceViewState) => void;
}>({
  name: 'wordSurfaceView',
  addProseMirrorPlugins() {
    const options = this.options;
    const editor = this.editor;
    return [
      new Plugin({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply: (tr, previous) => tr.getMeta(key) ?? previous.map(tr.mapping, tr.doc),
        },
        props: { decorations: (state) => key.getState(state) },
        view(view) {
          const storyMeasurer = wordStoryMeasurer(() => schedule());
          let frame = 0,
            destroyed = false,
            decorationSignature = '',
            outputSignature = '',
            reflowSignature = '';
          const media = window.matchMedia('(min-width: 701px)');
          const print = window.matchMedia('print');
          let printable:
            { input: SurfaceInput; plan: WordFragmentPlan; doc: typeof view.state.doc } | undefined;
          let printTree: HTMLElement | undefined;
          let pdfPrintable: typeof printable;
          pdfLayouts.set(view, (content) => {
            if (!pdfPrintable || pdfPrintable.doc !== view.state.doc)
              throw Error(
                'PDF export is not yet available for this layout. Use Print / Save as PDF.',
              );
            return measureWordPdf(view, pdfPrintable.input, pdfPrintable.plan, content);
          });
          const afterPrint = () => {
            printTree?.remove();
            printTree = undefined;
            view.dom.classList.remove('word-fragment-print-source');
          };
          const beforePrint = () => {
            afterPrint();
            if (!printable || printable.doc !== view.state.doc) return;
            printTree = wordFragmentPrint(view, printable.input, printable.plan);
            view.dom.after(printTree);
            view.dom.classList.add('word-fragment-print-source');
            alignWordPrintGlyphs(printTree);
          };
          const publish = (state: SurfaceViewState) => {
            const signature = JSON.stringify(state);
            if (signature !== outputSignature) {
              outputSignature = signature;
              options.change(state);
            }
          };
          const decorate = (
            specs: { from: number; to: number; attrs: Record<string, string>; inline?: boolean }[],
          ) => {
            const signature = JSON.stringify(specs);
            if (signature === decorationSignature) return;
            decorationSignature = signature;
            syncWordSelection(editor);
            view.dispatch(
              view.state.tr
                .setMeta(
                  key,
                  DecorationSet.create(
                    view.state.doc,
                    specs.map((s) =>
                      s.inline
                        ? Decoration.inline(s.from, s.to, s.attrs)
                        : Decoration.node(s.from, s.to, s.attrs),
                    ),
                  ),
                )
                .setMeta('addToHistory', false)
                .setMeta('preventUpdate', true),
            );
          };
          const measure = () => {
            frame = 0;
            if (destroyed || view.composing || print.matches) return;
            printable = undefined;
            pdfPrintable = undefined;
            exportLayouts.delete(view);
            pendingStories.delete(view);
            try {
              const content: WordContent = {
                ...options.content(),
                ...view.state.doc.attrs.wordPageLayout,
                sectionState: view.state.doc.attrs.wordSectionState || undefined,
                ...(view.state.doc.attrs.wordStories
                  ? { stories: view.state.doc.attrs.wordStories }
                  : {}),
              };
              const authored = authoredWordSection(content);
              const sections = authored ? [authored] : resolveWordSections(content);
              if (!sections.length || !media.matches) {
                decorate([]);
                publish(empty);
                return;
              }
              const hasStories = sections.some(
                (s) =>
                  Object.values(s.headers).some(Boolean) || Object.values(s.footers).some(Boolean),
              );
              const stories =
                hasStories && content.stories
                  ? storyMeasurer.read(content.stories, sections, content.docxStructure?.compatibility)
                  : null;
              if (hasStories && content.stories && stories === undefined) {
                pendingStories.add(view);
                decorate([]);
                publish({ ...empty, reason: 'Preparing headers and footers…' });
                return;
              }
              const measuredTabParagraphs = new Set<number>();
              if (sections.length === 1 && !sections[0].columns) view.state.doc.forEach((node, from) => {
                if (node.type.name !== 'paragraph') return;
                const paragraph = view.nodeDOM(from);
                if (!(paragraph instanceof HTMLElement)) return;
                const tabs = [...paragraph.querySelectorAll<HTMLElement>('[data-word-tab]')].map(wordTabPaintBox);
                if (!tabs.length || tabs.some((tab) => tab.dataset.wordTabMeasured !== 'true'
                  || (tab.dataset.wordTabLeader !== 'none' && tab.dataset.wordTabLeaderPainted !== 'true'))) return;
                const style = getComputedStyle(paragraph), height = parseFloat(style.height), line = parseFloat(style.lineHeight);
                if (line > 0 && Number.isFinite(height) && Math.abs(height - line) <= .02)
                  measuredTabParagraphs.add(from);
              });
              const input = content.docxStructure?.sections.some(
                (s) => s.drawings.length || s.notes.length,
              )
                ? null
                : sectionSurfaceInput(
                    view.state.doc,
                    authored
                      ? authoredWordSectionMap(view.state.doc, authored)
                      : wordSectionMap(view.state),
                    sections,
                    stories?.measurements,
                    measuredTabParagraphs,
                    content.numbering,
                  );
              if (!input) {
                decorate([]);
                publish({
                  ...empty,
                  reason:
                    hasStories && !stories
                      ? 'Continuous view: these headers or footers need additional layout support.'
                      : 'Continuous view: this section layout is not yet paginated.',
                });
                return;
              }
              if (input.blocks.some(b => {
                if (!b.list) return false;
                const dom = view.nodeDOM(b.from);
                if (!(dom instanceof HTMLElement) || !dom.hasAttribute('data-word-list-baseline')) return true;
                if (!b.list.marker) return false;
                const width = parseFloat(getComputedStyle(dom, '::before').width);
                return !Number.isFinite(width) || width >= b.list.definition.hanging / 15;
              })) {
                decorate([]);
                publish({ ...empty, reason: 'Continuous view: these list markers need additional layout support.' });
                return;
              }
              if (stories) {
                input.storyHtml = stories.html;
                input.storyText = stories.text;
              }
              // Measure actual editable blocks at their section body width. ProseMirror
              // owns every attribute change through decorations, including their removal.
              const widths = input.blocks.map((b) => ({
                from: b.from,
                to: b.to,
                attrs: {
                  class:
                    'word-measured-block' +
                    (b.flowBreaks ? ' word-flow-breaks' : '') +
                    (b.pageBreakParagraph ? ' word-page-break-paragraph' : ''),
                  style: `--surface-width:${b.width}px`,
                },
              }));
              // Existing planned blocks already have the correct measurement width.
              const ready = input.blocks.every((b) => {
                const dom = view.nodeDOM(b.from);
                return (
                  dom instanceof HTMLElement &&
                  Math.abs(parseFloat(getComputedStyle(dom).width) - b.width) < 0.1 &&
                  dom.classList.contains('word-flow-breaks') === !!b.flowBreaks &&
                  dom.classList.contains('word-page-break-paragraph') === !!b.pageBreakParagraph
                );
              });
              if (!ready || view.dom.querySelector('.word-page-fragment')) decorate(widths);
              const measurements = input.blocks.map((b) => {
                const dom = view.nodeDOM(b.from);
                if (!(dom instanceof HTMLElement)) throw Error('Missing paragraph surface.');
                const style = getComputedStyle(dom);
                return {
                  height: parseFloat(style.height),
                  before: parseFloat(style.marginTop),
                  after: parseFloat(style.marginBottom),
                };
              });
              let plan: WordSurfacePlan | WordFragmentPlan | null = planSectionSurfaces(
                input,
                measurements,
              );
              let reflow: WordColumnReflow | null = null;
              const providers = new Map<number, WordLineProvider>();
              const hasBarStops = input.blocks.some((b) =>
                wordTabBarOffsets(view.state.doc.nodeAt(b.from)?.attrs.paragraphTabs).length > 0);
              const hasLeaders = !!view.dom.querySelector('[data-word-tab-leader-painted="true"]');
              const measuredBaselines = input.blocks.every((b) => {
                const node = view.state.doc.nodeAt(b.from), dom = view.nodeDOM(b.from);
                return dom instanceof HTMLElement && dom.hasAttribute('data-word-measured-leading') &&
                  (node?.content.size === 0 || [...dom.querySelectorAll<HTMLElement>('[data-word-native-baseline]')]
                    .some((run) => Number.isFinite(Number(run.dataset.wordNativeBaseline))));
              });
              if (plan) {
                pdfPrintable = {
                  input,
                  plan: fragmentWordSurfaces(plan, measurements),
                  doc: view.state.doc,
                };
                // Short measured paragraphs need the same physical origins as
                // direct PDF. Browser flow rounds each paragraph's paint origin
                // independently, losing fractional Single-line advances.
                if (hasBarStops || hasLeaders || measuredBaselines || input.blocks.some(b => b.list)) printable = pdfPrintable;
              }
              // A single short section already uses its exact source rectangle.
              // Retain that surface; introduce page fragments only on overflow.
              if (plan && sections.length === 1 && !input.blocks.some(b => b.list)) {
                const grid = wordBodyGridOffsets(view, input, pdfPrintable!.plan);
                decorate([...grid].map(([index, shift]) => ({
                  from: input.blocks[index].from, to: input.blocks[index].to,
                  attrs: { 'data-word-body-grid': 'true', style: `translate:0 ${shift}px` },
                })));
                // The source screen rectangle stays in place, but the print
                // clone owns its margins and therefore needs its named pages.
                publish(printable ? { ...empty, printRules: sectionPrintRules(input) } : empty);
                return;
              }
              if (!plan && input.blocks.reduce((n, b) => n + b.to - b.from - 2, 0) <= 50000) {
                const flows = input.blocks.map((b, i) => {
                  const node = view.state.doc.nodeAt(b.from)!;
                  // Measured variable lines carry their own semantic ranges and
                  // advances. Unmeasured automatic/minimum text remains guarded.
                  const dom = view.nodeDOM(b.from);
                  const measured =
                    dom instanceof HTMLElement && dom.hasAttribute('data-word-measured-leading');
                  const lines =
                    measured ||
                    (node.attrs.paragraphLineRule !== 'atLeast' &&
                      /(?:pt|px)$/.test(node.attrs.paragraphLineHeight || ''))
                      ? measureWordLines(view, b.from)
                      : null;
                  const columns = sections[b.section].columns;
                  if (
                    lines &&
                    columns?.widths.some((w) => Math.abs(w - columns.widths[0]) > 0.001)
                  ) {
                    const text = paragraphGraphemes(node)?.text;
                    if (
                      text === undefined ||
                      lines.some(
                        (line, index) =>
                          index + 1 < lines.length &&
                          !['\n', '\f', '\u000e'].includes(text[line.to - 1]),
                      )
                    ) {
                      const provider = wordColumnLineProvider(view, b.from, lines);
                      if (!provider) return null;
                      providers.set(i, provider);
                    }
                  }
                  return lines
                    ? {
                        ...measurements[i],
                        height: lines.at(-1)!.top + lines.at(-1)!.height,
                        lines,
                        keepLines: node.attrs.keepLines === true,
                        widowControl: node.attrs.widowControl !== false,
                      }
                    : null;
                });
                if (flows.every((f) => f !== null)) {
                  if (providers.size) {
                    try {
                      reflow = reflowWordColumns(input, flows, providers);
                    } finally {
                      for (const provider of providers.values()) provider.dispose?.();
                    }
                    plan = reflow?.plan ?? null;
                  } else plan = planWordFragments(input, flows);
                }
              }
              if (!plan) {
                // Retain section widths for stable overflow measurement and recovery.
                decorate(widths);
                publish({
                  ...empty,
                  reason: 'Continuous view: this content needs paragraph pagination.',
                });
                return;
              }
              plan.tabBars = wordSurfaceTabBars(view, input,
                'fragments' in plan ? plan : fragmentWordSurfaces(plan, measurements));
              const specs: {
                from: number;
                to: number;
                attrs: Record<string, string>;
                inline?: boolean;
              }[] = plan.blocks.map((b) => ({
                from: b.from,
                to: b.to,
                attrs: {
                  class:
                    'word-measured-block word-surface-block' +
                    (b.flowBreaks ? ' word-flow-breaks' : '') +
                    (b.pageBreakParagraph ? ' word-page-break-paragraph' : ''),
                  'data-surface-section': String(b.section),
                  style: `--surface-width:${b.width}px;--surface-left:${b.left}px;--surface-top:${b.top}px;page:nofficeSection${sectionPrintPage(input, b.section)}`,
                },
              }));
              if ('fragments' in plan)
                for (const fragment of plan.fragments) {
                  if (fragment.to === fragment.from) continue;
                  const b = input.blocks[fragment.block];
                  const lines =
                    providers.has(fragment.block) && reflow
                      ? reflow.flows[fragment.block].lines.filter(
                          (line) => line.from >= fragment.from && line.to <= fragment.to,
                        )
                      : null;
                  for (const part of lines ?? [fragment]) {
                    const terminalBreak =
                      lines && view.state.doc.nodeAt(b.from + part.to)?.type.name === 'hardBreak';
                    specs.push({
                      from: b.from + 1 + part.from,
                      to: b.from + 1 + part.to - (terminalBreak ? 1 : 0),
                      inline: true,
                      attrs: {
                        nodeName: 'span',
                        class: 'word-page-fragment' + (lines ? ' word-reflow-line' : ''),
                        'data-fragment-page': String(fragment.page + 1),
                        'data-fragment-column': String(fragment.column + 1),
                        ...(lines ? { 'data-word-reflow-top': String(part.top) } : {}),
                        style:
                          `--fragment-shift:${fragment.shift}px;--fragment-left:${fragment.shiftX}px` +
                          (lines
                            ? `;display:inline-block;vertical-align:top;white-space:pre;width:${input.sections[b.section].columns!.widths[fragment.column] / 15}px;height:${part.height}px;line-height:${part.height}px`
                            : ''),
                      },
                    });
                    // A hard break is a line terminator, not another fixed-height
                    // line box. ProseMirror splits decorations at mark boundaries;
                    // an unmarked BR must not acquire a second inline-block row.
                    if (terminalBreak)
                      specs.push({
                        from: b.from + part.to,
                        to: b.from + 1 + part.to,
                        inline: true,
                        attrs: {
                          nodeName: 'span',
                          class: 'word-page-fragment word-reflow-break',
                          'data-fragment-page': String(fragment.page + 1),
                          'data-fragment-column': String(fragment.column + 1),
                          style: `display:inline;--fragment-shift:${fragment.shift}px;--fragment-left:${fragment.shiftX}px`,
                        },
                      });
                  }
                }
              decorate(specs);
              if ('fragments' in plan) pdfPrintable = { input, plan, doc: view.state.doc };
              const nextReflow = JSON.stringify(
                reflow ? [...providers.keys()].map((block) => reflow!.flows[block].lines) : [],
              );
              if (nextReflow !== reflowSignature) {
                const changed = reflowSignature !== '' || !!reflow;
                reflowSignature = nextReflow;
                if (changed) view.dom.dispatchEvent(new Event(wordLayoutChangeEvent));
              }
              if (
                'fragments' in plan &&
                (hasBarStops || hasLeaders || measuredBaselines || input.blocks.some(b => b.list) || input.stories ||
                  input.sections.some(
                    (s) => s.columns || ['oddPage', 'evenPage'].includes(s.start),
                  ) ||
                  input.blocks.some((b) => b.flowBreaks))
              )
                printable = { input, plan, doc: view.state.doc };
              if ('fragments' in plan && providers.size) {
                const paragraphs = input.blocks.map((b, block) => {
                  const node = view.state.doc.nodeAt(b.from)!;
                  const text = node.textBetween(0, node.content.size, '', (leaf) =>
                    leaf.type.name === 'wordPageBreak'
                      ? '\f'
                      : leaf.type.name === 'wordColumnBreak'
                        ? '\u000e'
                        : leaf.type.name === 'wordTab'
                          ? '\t'
                          : leaf.type.name === 'wordHyphen'
                            ? wordHyphenText(leaf.attrs.kind, 'native') || ''
                          : '\n',
                  );
                  const boundaries = [
                    ...new Set(
                      plan.fragments
                        .filter(
                          (f) =>
                            f.block === block &&
                            f.from > 0 &&
                            f.from < text.length &&
                            !['\n', '\f', '\u000e'].includes(text[f.from - 1]),
                        )
                        .map((f) => f.from),
                    ),
                  ].sort((a, b) => a - b);
                  return { text, boundaries };
                });
                exportLayouts.set(view, { doc: view.state.doc, paragraphs });
              }
              publish({
                plan,
                printRules: sectionPrintRules(input),
                reason: '',
                ...(stories ? { storyHtml: stories.html } : {}),
              });
            } catch {
              decorate([]);
              publish({
                ...empty,
                reason: 'Continuous view: section layout could not be measured.',
              });
            }
          };
          const schedule = () => {
            if (!destroyed && !frame) frame = requestAnimationFrame(measure);
          };
          const printChanged = () => {
            // Chromium's beforeprint can precede activation of print CSS.
            // Once the physical print grid is active, remeasure the temporary
            // clone at its final zoom rather than retaining screen offsets.
            if (print.matches) {
              if (printTree) alignWordPrintGlyphs(printTree);
            } else schedule();
          };
          const observer = new ResizeObserver(schedule);
          observer.observe(view.dom);
          document.fonts.addEventListener('loadingdone', schedule);
          view.dom.addEventListener('compositionend', schedule);
          view.dom.addEventListener(wordLayoutChangeEvent, schedule);
          media.addEventListener('change', schedule);
          print.addEventListener('change', printChanged);
          window.addEventListener('beforeprint', beforePrint);
          window.addEventListener('afterprint', afterPrint);
          document.fonts.ready.then(schedule);
          schedule();
          return {
            update(current, previous) {
              // A replacement can remove decorations even when all positions and
              // measurements remain equal (for example replacing selected text).
              if (current.state.doc !== previous.doc) {
                decorationSignature = '';
                schedule();
              } else if (
                wordSectionMap(current.state) !== wordSectionMap(previous) ||
                wordFontLineMetricsKey.getState(current.state) !==
                  wordFontLineMetricsKey.getState(previous) ||
                wordTabLayoutKey.getState(current.state) !== wordTabLayoutKey.getState(previous)
              ) {
                // Leading decorations settle after document edits. Absolutely
                // positioned blocks can grow without resizing the editor itself.
                schedule();
              }
            },
            destroy() {
              storyMeasurer.destroy();
              pendingStories.delete(view);
              pdfLayouts.delete(view);
              destroyed = true;
              cancelAnimationFrame(frame);
              observer.disconnect();
              document.fonts.removeEventListener('loadingdone', schedule);
              view.dom.removeEventListener('compositionend', schedule);
              view.dom.removeEventListener(wordLayoutChangeEvent, schedule);
              media.removeEventListener('change', schedule);
              print.removeEventListener('change', printChanged);
              window.removeEventListener('beforeprint', beforePrint);
              window.removeEventListener('afterprint', afterPrint);
              afterPrint();
            },
          };
        },
      }),
    ];
  },
});
