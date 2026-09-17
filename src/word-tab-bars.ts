import { wordTabStops } from './word-tab-stops';
import type { EditorView } from '@tiptap/pm/view';
import type { SurfaceInput } from './word-section-surfaces';
import type { WordFragmentPlan } from './word-fragment-plan';

/** Native bars paint a black hairline immediately inside the stop position,
 * across the line box, independently of text color and literal tab characters.
 * CSS pixels here; the measured 0.14pt native stroke is represented by 3 twips. */
export const wordTabBarWidth = 0.2;

export function wordTabBarOffsets(value: unknown, indent = 0) {
  return (wordTabStops(value) || [])
    .filter((stop) => stop.alignment === 'bar')
    .map((stop) => stop.position / 15 - indent);
}

/** Screen page backgrounds need separate rectangles: the editable paragraph's
 * inline fragments can occupy multiple columns/pages while its box stays put. */
export function wordSurfaceTabBars(view: EditorView, input: SurfaceInput, plan: WordFragmentPlan) {
  const result: NonNullable<WordFragmentPlan['tabBars']> = [];
  for (const fragment of plan.fragments) {
    const node = view.state.doc.nodeAt(input.blocks[fragment.block].from);
    if (!node || node.attrs.paragraphTabUnsupported || node.attrs.direction === 'rtl') continue;
    for (const offset of wordTabBarOffsets(node.attrs.paragraphTabs)) {
      if (result.length >= 5000) throw Error('The tab bar layout exceeds the drawing limit.');
      result.push({ page: fragment.page, left: fragment.left + offset, top: fragment.top,
        width: wordTabBarWidth, height: fragment.height });
    }
  }
  return result;
}
