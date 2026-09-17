import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { wordHyphenText } from './word-hyphen';
import { wordTabLayoutKey } from './word-tab-state';
import { WordFontLineMetrics, wordFontLineMetricsKey } from './word-font-line-metrics';
import { measureWordLines } from './word-line-measurements';
import { wordStoryParagraphPaint } from './word-story-paint';
import { wordStoryGridEligible } from './word-story-grid';
import { WordJustification } from './word-justification';
import { wordJustificationKey } from './word-justification-state';
import type { WordCompatibility } from './word-compatibility';
import { sanitizeHTML } from './formats';
import type { WordStories } from './word-stories';
import type { WordSectionLayout } from './word-section-layout';
import type { WordStoryMeasurements } from './word-story-layout';

export interface WordStoryRendering {
  measurements: WordStoryMeasurements;
  html: Record<string, string>;
  text: Record<string, string>;
}

/** Cache story reading views by source part, section width and physical left
 * margin. Leader grids depend on page origin even at equal widths. Bound
 * the derived views, and discard stale asynchronous measurements on every edit. */
export function wordStoryMeasurer(changed: () => void) {
  let signature = '',
    frame = 0,
    generation = 0;
  let views: {
    part: WordStories['parts'][number];
    width: number;
    pageLeft: number;
    key: string;
    editor: Editor;
    host: HTMLElement;
  }[] = [];
  let result: WordStoryRendering | null | undefined;
  const clear = () => {
    generation++;
    cancelAnimationFrame(frame);
    for (const { editor, host } of views) {
      editor.destroy();
      host.remove();
    }
    views = [];
    result = undefined;
  };
  const invalidate = () => {
    signature = '';
    clear();
    changed();
  };
  document.fonts.addEventListener('loadingdone', invalidate);
  return {
    read(
      stories: WordStories,
      sections: WordSectionLayout[],
      compatibility?: WordCompatibility,
    ): WordStoryRendering | null | undefined {
      const next = JSON.stringify([stories, sections, compatibility]);
      if (next === signature) return result;
      clear();
      signature = next;
      const needed = new Map<
        string,
        { part: WordStories['parts'][number]; width: number; pageLeft: number; key: string }
      >();
      if (!sections.length || sections.length > 100) return (result = null);
      for (const section of sections) {
        const width =
          ((section.width ?? 0) - (section.margins.left ?? 0) - (section.margins.right ?? 0)) / 15;
        const pageLeft = section.margins.left === null ? NaN : section.margins.left / 15;
        if (!(width > 0) || !Number.isFinite(pageLeft) || pageLeft < 0) return (result = null);
        for (const kind of ['headers', 'footers'] as const)
          for (const slot of ['default', 'first', 'even'] as const) {
            if (
              (slot === 'first' && !section.differentFirstPage) ||
              (slot === 'even' && !stories.evenAndOddHeaders)
            )
              continue;
            const ref = section[kind][slot];
            if (!ref) continue;
            const parts = stories.parts.filter((p) =>
              p.relationshipIds.includes(ref.relationshipId),
            );
            if (parts.length !== 1) return (result = null);
            const part = parts[0];
            const key = JSON.stringify([part.path, width, pageLeft]);
            needed.set(key, { part, width, pageLeft, key });
            // Prevent a hostile many-section document from constructing an
            // unbounded number of complete editor/plugin trees synchronously.
            if (needed.size > 64) return (result = null);
          }
      }
      for (const { part, width, pageLeft, key } of needed.values()) {
        const host = document.createElement('div');
        host.className = 'word-story-measure-host';
        host.setAttribute('aria-hidden', 'true');
        host.style.cssText =
          'position:fixed;left:-100000px;top:0;visibility:hidden;pointer-events:none';
        document.body.append(host);
        try {
          const editor = new Editor({
            element: host,
            editable: false,
            extensions: [WordJustification.configure({ compatibility: () => compatibility }), WordFontLineMetrics, ...wordExtensions({ tabPageLeft: () => pageLeft })],
            content: sanitizeHTML(part.html),
            parseOptions: { preserveWhitespace: true },
            editorProps: {
              attributes: {
                class: 'document-page',
                style: `padding:0!important;min-height:0!important;width:${width}px;height:auto;overflow:visible`,
              },
            },
          });
          views.push({ part, width, pageLeft, key, editor, host });
        } catch {
          host.remove();
          clear();
          signature = next;
          return (result = null);
        }
      }
      let previous: unknown[] = [],
        stable = 0,
        passes = 0;
      const own = generation;
      const measure = () => {
        frame = 0;
        if (own !== generation) return;
        try {
          const current = views.flatMap(({ editor }) => [
            wordFontLineMetricsKey.getState(editor.state),
            wordTabLayoutKey.getState(editor.state),
            wordJustificationKey.getState(editor.state),
          ]);
          stable =
            document.fonts.status === 'loaded' && current.every((value, i) => value === previous[i])
              ? stable + 1
              : 0;
          previous = current;
          if (stable < 3 && ++passes < 60) {
            frame = requestAnimationFrame(measure);
            return;
          }
          if (stable < 3) {
            result = null;
            changed();
            return;
          }
          const measured: WordStoryMeasurements['parts'] = [];
          const html: Record<string, string> = Object.create(null);
          const text: Record<string, string> = Object.create(null);
          let valid = true;
          for (const { part, width, pageLeft, key, editor } of views) {
            let height = 0, previousAfter = 0, gridEligible = true, variableGridEligible = true, spaced = false;
            const copy = document.createElement('div');
            editor.state.doc.forEach((node, from) => {
              if (
                node.type.name !== 'paragraph' ||
                node.attrs.direction === 'rtl'
              ) {
                valid = false;
                return;
              }
              const element = editor.view.nodeDOM(from);
              if (!(element instanceof HTMLElement)) {
                valid = false;
                return;
              }
              const style = getComputedStyle(element),
                lines = measureWordLines(editor.view, from),
                before = parseFloat(style.marginTop), after = parseFloat(style.marginBottom),
                left = parseFloat(style.marginInlineStart), right = parseFloat(style.marginInlineEnd),
                first = parseFloat(style.textIndent), availableWidth = width - left - right;
              if (
                !lines?.length ||
                ![before, after].every((value) => Number.isFinite(value) && value >= 0 && value <= 2112) ||
                ![left, right, first].every(value => Number.isFinite(value) && Math.abs(value) <= 960) ||
                !(availableWidth > 0) || first >= availableWidth ||
                lines.some((line) => line.breakAfter) ||
                !(parseFloat(style.height) > 0)
              ) {
                valid = false;
                return;
              }
              // Pinned header/footer probes retain leading/trailing spacing
              // and collapse adjacent after/before values to their maximum.
              height += Math.max(previousAfter, before);
              gridEligible &&= wordStoryGridEligible(element, lines.length);
              variableGridEligible &&= wordStoryGridEligible(element, lines.length, true);
              spaced ||= before > 0 || after > 0;
              const paints = wordStoryParagraphPaint(editor.view, from, lines);
              for (const [index, { element: clone, top }] of paints.entries()) {
                // Wrapped line clones are still one semantic paragraph.
                if (index > 0) clone.style.marginTop = '0';
                if (index < paints.length - 1) clone.style.marginBottom = '0';
                clone.dataset.wordStoryParagraphTop = String(height + top);
                clone.dataset.wordStoryGridAfter = String(Math.round(after * 15) / 15);
                clone.dataset.wordStoryParagraphLeft = String(left);
                clone.dataset.wordStoryParagraphWidth = String(availableWidth);
                clone.style.width = `${availableWidth}px`;
                copy.append(clone);
              }
              height += lines.at(-1)!.top + lines.at(-1)!.height;
              previousAfter = after;
            });
            height += previousAfter;
            // Native mixed line-spacing stories quantize each paragraph origin
            // independently. With zero Before/After the existing grid formula
            // reduces to origin rounding, regardless of the physical line height.
            // Keep the established all-exact40 zero-spacing path unchanged.
            const variableZeroSpacing = !spaced && !gridEligible && variableGridEligible;
            if (!(gridEligible && spaced) && !variableZeroSpacing)
              for (const p of copy.querySelectorAll<HTMLElement>('[data-word-story-grid-after]'))
                delete p.dataset.wordStoryGridAfter;
            if (!(height > 0) || !Number.isFinite(height)) valid = false;
            measured.push({
              path: part.path,
              kind: part.kind,
              relationshipIds: part.relationshipIds,
              height,
              width,
              pageLeft,
              renderKey: key,
            });
            html[key] = copy.innerHTML;
            text[key] = editor.state.doc.textBetween(0, editor.state.doc.content.size, '',
              (leaf) => leaf.type.name === 'wordTab' ? '\t' : leaf.type.name === 'wordHyphen' ? wordHyphenText(leaf.attrs.kind) || '' : '');
          }
          result = valid
            ? {
                measurements: {
                  version: 1,
                  evenAndOddHeaders: stories.evenAndOddHeaders,
                  parts: measured,
                },
                html,
                text,
              }
            : null;
          changed();
        } catch {
          result = null;
          changed();
        }
      };
      frame = requestAnimationFrame(measure);
      return result;
    },
    destroy() {
      document.fonts.removeEventListener('loadingdone', invalidate);
      clear();
    },
  };
}
