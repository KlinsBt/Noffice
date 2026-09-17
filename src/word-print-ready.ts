import type { EditorView } from '@tiptap/pm/view';
import { wordFontLineMetricsKey, wordLayoutChangeEvent } from './word-font-line-metrics';
import { wordSurfaceViewKey, wordStoryLayoutPending } from './word-surface-view';
import { wordTabLayoutKey } from './word-tab-state';

/** Wait for the current document's fonts and derived layout, not an arbitrary
 * delay. A changed document/layout restarts settling; nothing is saved or edited. */
export async function prepareWordPrint(view: EditorView) {
  const dom = view.dom;
  const window = dom.ownerDocument.defaultView!;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let frame = 0;
  let cancelled = false;
  const ready = async () => {
    await dom.ownerDocument.fonts.ready;
    if (cancelled) return;
    if (view.isDestroyed) throw Error('The document was closed before printing.');
    if (view.composing) throw Error('Finish entering the current text before printing.');
    dom.dispatchEvent(new Event(wordLayoutChangeEvent));
    let previous: unknown[] = [],
      stable = 0;
    while (stable < 3) {
      await new Promise<void>((resolve) => {
        frame = window.requestAnimationFrame(() => resolve());
      });
      if (cancelled) return;
      if (view.isDestroyed) throw Error('The document was closed before printing.');
      const current = [
        view.state.doc,
        wordFontLineMetricsKey.getState(view.state),
        wordTabLayoutKey.getState(view.state),
        wordSurfaceViewKey.getState(view.state),
        dom.getBoundingClientRect().width,
      ];
      stable =
        !view.composing &&
        !wordStoryLayoutPending(view) &&
        dom.ownerDocument.fonts.status === 'loaded' &&
        current.every((value, i) => value === previous[i])
          ? stable + 1
          : 0;
      previous = current;
    }
  };
  try {
    await Promise.race([
      ready(),
      new Promise<never>((_, reject) => {
        timer = setTimeout(
          () =>
            reject(
              Error('Document layout is still changing. Try printing again after it settles.'),
            ),
          5000,
        );
      }),
    ]);
  } finally {
    cancelled = true;
    if (timer !== undefined) clearTimeout(timer);
    window.cancelAnimationFrame(frame);
  }
}
