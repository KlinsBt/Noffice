import type { EditorView } from '@tiptap/pm/view';
import type { SurfaceInput } from './word-section-surfaces';
import type { WordFragmentPlan } from './word-fragment-plan';
import { wordStoryGridEligible } from './word-story-grid';

/** Native body shading includes Before in the paragraph box. Round the full
 * box and its trailing space separately, preserving the semantic flow height.
 * Twelve independent source/shaded PDF+EMF controls establish this contract. */
export function wordBodyGridOrigin(top: number, before: number, after: number): number {
  if (![top, before, after].every(Number.isFinite) || top < before || before < 0 || after < 0 ||
      before > 801 / 15 || after > 801 / 15) return top;
  const b = Math.round(before * 15), a = Math.round(after * 15);
  return (Math.round((Math.round(top * 15) - b) * 5 / 12)
    + Math.round((800 + b + a) * 5 / 12) - Math.round(a * 5 / 12)
    - Math.round(800 * 5 / 12)) * .16;
}

/** Confine this measured paint contract to one short, regular Arial10/exact40
 * body. Wrapped lines, adjacent nonzero margins, columns and overflow retain
 * their existing path until their complete boundary matrix is established. */
export function wordBodyGridOffsets(view: EditorView, input: SurfaceInput, plan: WordFragmentPlan): Map<number, number> {
  const offsets = new Map<number, number>();
  if (input.sections.length !== 1 || input.sections[0].columns || plan.pages.length !== 1 ||
      plan.fragments.length !== input.blocks.length) return offsets;
  const values: { top: number; before: number; after: number; block: number }[] = [];
  for (const fragment of plan.fragments) {
    const block = input.blocks[fragment.block], dom = view.nodeDOM(block.from);
    if (!(dom instanceof HTMLElement) || !wordStoryGridEligible(dom, 1) || block.flowBreaks ||
        Math.abs(parseFloat(getComputedStyle(dom).height) - 800 / 15) > .02) return offsets;
    const css = getComputedStyle(dom), before = parseFloat(css.marginTop), after = parseFloat(css.marginBottom);
    if (before > 0 && (values.at(-1)?.after || 0) > 0) return offsets;
    values.push({ block: fragment.block, top: fragment.top - plan.pages[0].top, before, after });
  }
  for (const value of values) offsets.set(value.block, wordBodyGridOrigin(value.top, value.before, value.after) - value.top);
  return offsets;
}
