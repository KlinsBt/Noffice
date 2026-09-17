import { planWordFragments, type ParagraphFlow, type WordFragmentPlan } from './word-fragment-plan';
import type { WordLine } from './word-line-measurements';
import type { SurfaceInput } from './word-section-surfaces';

/** A measurement provider consumes a complete semantic line at the available
 * width. It must preserve grapheme boundaries and include its trailing spaces.
 * The allocator remains independent of fonts, DOM measurement and Svelte. */
export type WordLineProvider = ((from: number, width: number) => Omit<WordLine, 'top'> | null) & {
  dispose?: () => void;
};

export interface WordColumnReflow {
  plan: WordFragmentPlan;
  flows: ParagraphFlow[];
}

/** Reconcile line measurement with the columns chosen by the existing
 * keep/widow/section allocator. Widths follow visual line ordinals, not old text
 * offsets: a wider column can consume more text without retaining the narrow
 * column's old wrap boundaries. Cycles and exhausted work return no plan. */
export function reflowWordColumns(
  input: SurfaceInput,
  initial: ParagraphFlow[],
  providers: ReadonlyMap<number, WordLineProvider>,
): WordColumnReflow | null {
  if (initial.length !== input.blocks.length) return null;
  let flows = initial;
  const seen = new Set<string>();
  let measured = 0;
  for (let pass = 0; pass < 32; pass++) {
    const plan = planWordFragments(input, flows);
    if (!plan) return null;
    const signature = JSON.stringify(flows.map((flow) => flow.lines));
    if (seen.has(signature)) return null;
    seen.add(signature);
    const next = flows.slice();
    for (const [block, measure] of providers) {
      const b = input.blocks[block],
        flow = flows[block];
      if (!b || !flow) return null;
      const section = input.sections[b.section];
      if (!section.columns) return null;
      const fragments = plan.fragments.filter((fragment) => fragment.block === block);
      const widths = flow.lines.map((line) => {
        const fragment = fragments.find(
          (f) => line.from >= f.from && (line.from < f.to || f.from === f.to),
        );
        return fragment ? section.columns!.widths[fragment.column] / 15 : NaN;
      });
      if (!widths.length || widths.some((width) => !Number.isFinite(width) || width <= 0))
        return null;
      const lines: WordLine[] = [];
      let from = 0,
        top = 0;
      const end = b.to - b.from - 2;
      while (from < end) {
        if (++measured > 50000) return null;
        const width = widths[Math.min(lines.length, widths.length - 1)];
        const line = measure(from, width);
        if (
          !line ||
          line.from !== from ||
          !Number.isSafeInteger(line.to) ||
          line.to <= from ||
          line.to > end ||
          !Number.isFinite(line.height) ||
          line.height <= 0
        )
          return null;
        lines.push({ ...line, top });
        from = line.to;
        top += line.height;
      }
      if (!lines.length) return null;
      next[block] = { ...flow, height: top, lines };
    }
    if (JSON.stringify(next.map((flow) => flow.lines)) === signature) return { plan, flows };
    flows = next;
  }
  return null;
}
