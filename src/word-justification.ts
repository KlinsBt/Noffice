import { Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import type { WordCompatibility } from './word-compatibility';
import { wordJustifiedLines } from './word-justified-lines';
import { wordJustificationKey } from './word-justification-state';
import { wordFontSizePixels, wordLayoutChangeEvent } from './word-font-line-metrics';
import { wordTabLeaderNormalRatio } from './word-tab-leader-metrics';
import { wordLineSpacing } from './word-line-spacing';
import { syncWordSelection } from './word-links';

/** Disposable soft breaks and space advances. Text, marks, history and retained
 * DOCX runs stay semantic. Independently measured Arial10 regular/bold and
 * Calibri11 regular with standard/contextual ligatures use the exact40 mode15
 * path; other profiles retain their current layout. */
export const WordJustification = Extension.create<{
  compatibility: () => WordCompatibility | undefined;
}>({
  name: 'wordJustification',
  addOptions: () => ({ compatibility: () => undefined }),
  addProseMirrorPlugins() {
    const compatibility = this.options.compatibility, editor = this.editor;
    return [new Plugin({
      key: wordJustificationKey,
      state: {
        init: () => ({ signature: '', decorations: DecorationSet.empty }),
        apply: (tr, prior) => tr.getMeta(wordJustificationKey) ?? (tr.docChanged
          ? { signature: '', decorations: DecorationSet.empty } : prior),
      },
      props: {
        decorations: (state) => wordJustificationKey.getState(state)?.decorations,
        handleDOMEvents: {
          compositionstart(view) {
            syncWordSelection(editor);
            const { from, to } = view.state.selection;
            if (from !== to && wordJustificationKey.getState(view.state)?.decorations
              .find(from, to).some(decoration => decoration.spec.wordJustificationBreak)) {
              // Let the browser own candidate text in a continuous editable
              // range. Disposable breaks can otherwise make its replacement
              // parse as deletion. Layout resumes on compositionend.
              view.dispatch(view.state.tr.setMeta(wordJustificationKey, {
                signature: '', decorations: DecorationSet.empty,
              }).setMeta('addToHistory', false));
            }
            return false;
          },
          beforeinput(view, event) {
            if (!event.cancelable || event.isComposing || view.composing
              || event.inputType !== 'insertText' || !event.data) return false;
            syncWordSelection(editor);
            const { from, to } = view.state.selection;
            if (from === to || !wordJustificationKey.getState(view.state)?.decorations
              .find(from, to).some(decoration => decoration.spec.wordJustificationBreak)) return false;
            // Chromium can merge the nested paint spans while replacing a range
            // across a disposable break. Its DOM change then parses as deletion,
            // dropping the first typed character. Apply this text input to the
            // semantic selection before that mutation, retaining input handlers.
            event.preventDefault();
            const text = event.data;
            const transaction = () => view.state.tr.insertText(text, from, to).scrollIntoView();
            if (!view.someProp('handleTextInput', handler => handler(view, from, to, text, transaction)))
              view.dispatch(transaction());
            return true;
          },
        },
      },
      view(view) {
        let frame = 0, disposed = false;
        const ratios = new Map<string, number>(), advances = new Map<string, Map<string, number>>();
        const canvas = document.createElement('canvas'), context = canvas.getContext('2d');
        const measure = () => {
          frame = 0;
          if (disposed || !view.dom.isConnected || view.composing) return;
          const current = compatibility(), specs: {
            from: number; to: number;
            lines: NonNullable<ReturnType<typeof wordJustifiedLines>>;
          }[] = [];
          let characters = 0, paragraphs = 0;
          if (context && current?.mode === 15 && current.wordPerfectJustification === false
            && document.fonts.status === 'loaded') view.state.doc.forEach((node, from) => {
            if (++paragraphs > 5000 || node.type.name !== 'paragraph'
              || node.attrs.direction === 'rtl' || !node.childCount) return;
            const spacing = wordLineSpacing(node.attrs.paragraphLineHeight, node.attrs.paragraphLineRule);
            if (spacing?.rule !== 'exact' || spacing.line !== 800) return;
            let allText = true;
            node.forEach((child) => { if (!child.isText) allText = false; });
            if (!allText || node.textContent.length > 4096
              || (characters += node.textContent.length) > 10000) return;
            const paragraph = view.nodeDOM(from);
            if (!(paragraph instanceof HTMLElement)) return;
            const style = getComputedStyle(paragraph);
            if (style.textAlign !== 'justify' || style.direction !== 'ltr'
              || parseFloat(style.textIndent) !== 0
              || parseFloat(style.paddingLeft) !== 0 || parseFloat(style.paddingRight) !== 0) return;
            const width = Math.round(parseFloat(style.width) * 15) / 15;
            let face: string | undefined, valid = true, calibri = false;
            const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
            while (walker.nextNode()) {
              if (!walker.currentNode.textContent || walker.currentNode.parentElement?.closest('.ProseMirror-widget')) continue;
              const element = walker.currentNode.parentElement!, run = getComputedStyle(element);
              const semantic = view.state.doc.nodeAt(view.posAtDOM(walker.currentNode, 0));
              const textStyle = semantic?.marks.find((mark) => mark.type.name === 'textStyle')?.attrs;
              const size = wordFontSizePixels(textStyle?.fontSize || node.attrs.paragraphFontSize);
              const family = run.fontFamily.split(',')[0].replace(/["']/g, '').trim().toLowerCase();
              const weight = ['bold', '700'].includes(run.fontWeight) ? '700' : '400';
              const features = textStyle?.wordFontFeatures ?? node.attrs.paragraphFontFeatures;
              const shaped = family === 'calibri' && weight === '400' && features === 3;
              const profile = `${family}:${weight}`;
              if ((!shaped && family !== 'arial') || size === null
                || Math.abs(size - (shaped ? 44 : 40) / 3) > .0001
                || !['normal', '400', 'bold', '700'].includes(run.fontWeight)
                || run.fontStyle !== 'normal' || !['normal', '100%'].includes(run.fontStretch)
                || run.fontKerning !== 'none' || run.fontVariantCaps !== 'normal'
                || (run.letterSpacing !== 'normal' && parseFloat(run.letterSpacing) !== 0)
                || semantic?.marks.some((mark) => ['subscript', 'superscript'].includes(mark.type.name))
                || ![0, 3].includes(features)
                || Math.abs(wordTabLeaderNormalRatio(run, ratios) - (shaped ? 1.2207 : 1.1499)) > .0001
                || (face !== undefined && face !== profile)) valid = false;
              face = profile; calibri = shaped;
            }
            const descriptor = calibri ? '400 1024px Calibri' : `${face?.split(':')[1]} 1024px Arial`;
            if (!valid || !face || !document.fonts.check(descriptor)) return;
            let metrics = advances.get(face);
            if (!metrics) {
              context.font = descriptor;
              context.fontKerning = 'none';
              metrics = new Map(calibri ? [] : Array.from({ length: 95 }, (_,index) => {
                const text = String.fromCharCode(32 + index);
                return [text, context.measureText(text).width / 1024 * (40 / 3)] as const;
              }));
              advances.set(face, metrics);
            }
            // Calibri ligatures change whole-word advances. Measure at the
            // independently checked neutral size, without per-glyph rounding.
            context.font = descriptor; context.fontKerning = 'none';
            const lines = wordJustifiedLines(node.textContent, width, (text) => {
              if (!calibri) return [...text].reduce((sum, character) => sum + (metrics!.get(character) ?? NaN), 0);
              const prior = metrics!.get(text);
              if (prior !== undefined) return prior;
              const advance = context.measureText(text).width / 1024 * (44 / 3);
              if (metrics!.size < 2048) metrics!.set(text, advance);
              return advance;
            });
            if (lines && lines.length > 1) specs.push({ from, to: from + node.nodeSize, lines });
          });
          const signature = JSON.stringify(specs);
          if (signature === wordJustificationKey.getState(view.state)?.signature) return;
          const decorations: Decoration[] = [];
          for (const spec of specs) {
            decorations.push(Decoration.node(spec.from, spec.to, {
              'data-word-justification': 'modern', style: 'white-space:pre;text-align-last:left',
            }));
            for (const [index, line] of spec.lines.entries()) {
              decorations.push(Decoration.inline(spec.from + 1 + line.from, spec.from + 1 + line.to, {
                'data-word-justification-run': 'true', style: `word-spacing:${line.wordSpacing}px`,
              }));
              if (index) decorations.push(Decoration.widget(spec.from + 1 + line.from, () => {
                const br = document.createElement('br');
                br.dataset.wordJustificationBreak = 'true'; br.setAttribute('aria-hidden', 'true');
                return br;
              }, { side: -2, ignoreSelection: true, wordJustificationBreak: true, key: `justify-${spec.from}-${line.from}` }));
            }
          }
          syncWordSelection(editor);
          view.dispatch(view.state.tr.setMeta(wordJustificationKey, {
            signature, decorations: DecorationSet.create(view.state.doc, decorations),
          }).setMeta('addToHistory', false));
          view.dom.dispatchEvent(new Event(wordLayoutChangeEvent));
        };
        const schedule = () => { if (!disposed && !frame) frame = requestAnimationFrame(measure); };
        const fontsChanged = () => { ratios.clear(); advances.clear(); schedule(); };
        const resize = new ResizeObserver(schedule); resize.observe(view.dom);
        document.fonts.addEventListener('loadingdone', fontsChanged);
        document.fonts.ready.then(schedule);
        view.dom.addEventListener('compositionend', schedule);
        view.dom.addEventListener(wordLayoutChangeEvent, schedule);
        schedule();
        return {
          update(current, previous) { if (current.state.doc !== previous.doc) schedule(); },
          destroy() {
            disposed = true; cancelAnimationFrame(frame); resize.disconnect();
            document.fonts.removeEventListener('loadingdone', fontsChanged);
            view.dom.removeEventListener('compositionend', schedule);
            view.dom.removeEventListener(wordLayoutChangeEvent, schedule);
          },
        };
      },
    })];
  },
});
