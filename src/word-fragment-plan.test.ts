import { expect, it } from 'vitest';
import { getSchema } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { paragraphGraphemes } from './word-line-measurements';
import { planWordFragments, type ParagraphFlow } from './word-fragment-plan';
import type { SurfaceInput } from './word-section-surfaces';

function fixture(
  counts: number[],
  keep = -1,
  widow = false,
): { input: SurfaceInput; metrics: ParagraphFlow[] } {
  let pos = 0;
  const input: SurfaceInput = {
    width: 200,
    sections: [
      {
        id: 's1',
        width: 3000,
        height: 3000,
        margins: { left: 300, right: 300, top: 300, bottom: 300, header: 0, footer: 0, gutter: 0 },
        orientation: 'portrait',
        paragraphs: [],
        headers: { default: null, even: null, first: null },
        footers: { default: null, even: null, first: null },
        simpleBody: true,
        differentFirstPage: false,
        start: 'nextPage',
      },
    ],
    blocks: counts.map((n) => {
      const from = pos;
      pos += n + 2;
      return { from, to: pos, section: 0, width: 160 };
    }),
  };
  const metrics = counts.map((n, i) => ({
    height: n * 20,
    before: 0,
    after: 0,
    keepLines: keep === i,
    widowControl: widow,
    lines: Array.from({ length: n }, (_, j) => ({ from: j, to: j + 1, top: j * 20, height: 20 })),
  }));
  return { input, metrics };
}
it('splits line fragments at page capacity without changing semantic paragraphs', () => {
  const { input, metrics } = fixture([12]);
  const before = JSON.stringify(input),
    plan = planWordFragments(input, metrics)!;
  expect(plan.pages).toHaveLength(2);
  expect(plan.blocks).toHaveLength(1);
  expect(plan.fragments.map((f) => [f.from, f.to, f.page, f.shift])).toEqual([
    [0, 8, 0, 0],
    [8, 12, 1, 64],
  ]);
  expect(JSON.stringify(input)).toBe(before);
});
it('moves a kept paragraph but splits one larger than a page', () => {
  const f = fixture([6, 4], 1),
    p = planWordFragments(f.input, f.metrics)!;
  expect(p.fragments.map((f) => [f.block, f.page])).toEqual([
    [0, 0],
    [1, 1],
  ]);
  const large = fixture([12], 0);
  expect(planWordFragments(large.input, large.metrics)!.pages).toHaveLength(2);
});
it('moves two orphan lines and pulls a widow line back with its predecessor', () => {
  const f = fixture([7, 3], -1, true);
  expect(
    planWordFragments(f.input, f.metrics)!.fragments.map((f) => [f.block, f.page, f.from, f.to]),
  ).toEqual([
    [0, 0, 0, 7],
    [1, 1, 0, 3],
  ]);
  const w = fixture([9], -1, true);
  expect(planWordFragments(w.input, w.metrics)!.fragments.map((f) => [f.from, f.to])).toEqual([
    [0, 7],
    [7, 9],
  ]);
});
it('fits kept paragraphs by their final natural box while preserving internal advances', () => {
  const f = fixture([5, 3], 1);
  f.metrics[0].height = 110;
  f.metrics[0].lines.forEach((line, i) => {
    line.height = 22;
    line.top = i * 22;
  });
  f.metrics[1].lines.forEach((line) => Object.assign(line, { fitHeight: 10 }));
  const plan = planWordFragments(f.input, f.metrics)!;
  expect(plan.pages).toHaveLength(1);
  expect(plan.fragments.map((f) => [f.block, f.top, f.height])).toEqual([
    [0, 20, 110],
    [1, 130, 60],
  ]);
});
it('keeps the following paragraph when its last text box fits beyond the preceding full advances', () => {
  const f = fixture([6, 2]);
  f.input.blocks[0].keepNext = true;
  f.metrics[0].after = 10;
  f.metrics[1].lines.forEach((line) => Object.assign(line, { fitHeight: 10 }));
  const plan = planWordFragments(f.input, f.metrics)!;
  expect(plan.pages).toHaveLength(1);
  expect(plan.fragments.map((f) => [f.block, f.from, f.to])).toEqual([
    [0, 0, 6],
    [1, 0, 2],
  ]);
});
it('rejects invalid occupied line heights instead of accepting impossible page fits', () => {
  for (const fitHeight of [0, -1, NaN, Infinity, 21]) {
    const f = fixture([3]);
    Object.assign(f.metrics[0].lines[0], { fitHeight });
    expect(planWordFragments(f.input, f.metrics)).toBeNull();
  }
});
it('collapses adjacent spacing and suppresses consumed spacing at an automatic page start', () => {
  const f = fixture([7, 2]);
  f.metrics[0].after = 10;
  f.metrics[1].before = 15;
  const plan = planWordFragments(f.input, f.metrics)!;
  expect(plan.blocks[1].top).toBe(224 + 20 - 15);
});
it('rejects an oversized line, discontinuous offsets, invalid metrics and excessive output', () => {
  for (const mutate of [
    (f: ReturnType<typeof fixture>) => {
      f.metrics[0].lines[0].height = 1000;
    },
    (f: ReturnType<typeof fixture>) => {
      f.metrics[0].lines[1].from = 3;
    },
    (f: ReturnType<typeof fixture>) => {
      f.metrics[0].before = NaN;
    },
  ]) {
    const f = fixture([3]);
    mutate(f);
    expect(planWordFragments(f.input, f.metrics)).toBeNull();
  }
  const f = fixture([5000]);
  expect(planWordFragments(f.input, f.metrics)).toBeNull();
});
it('segments full paragraph graphemes across marks without splitting surrogate pairs or combining sequences', () => {
  const schema = getSchema(wordExtensions());
  const p = schema.nodeFromJSON({
    type: 'paragraph',
    content: [
      { type: 'text', text: 'A\ud83d\ude00e' },
      { type: 'text', text: '\u0301', marks: [{ type: 'bold' }] },
      { type: 'hardBreak' },
      { type: 'text', text: 'Z' },
    ],
  });
  expect(paragraphGraphemes(p)).toEqual({
    text: 'A\ud83d\ude00e\u0301\nZ',
    boundaries: [0, 1, 3, 5, 6, 7],
  });
  expect(
    paragraphGraphemes(schema.nodeFromJSON({ type: 'paragraph', content: [{ type: 'wordTab' }] })),
  ).toEqual({ text: '\t', boundaries: [0, 1] });
});
it('anchors a terminal empty line to the real paragraph and shifts earlier text back', () => {
  const f = fixture([9]);
  f.metrics[0].lines[8].to = 8;
  f.input.blocks[0].to--;
  const plan = planWordFragments(f.input, f.metrics)!;
  expect(plan.blocks[0].top).toBe(84);
  expect(plan.fragments.map((f) => [f.from, f.to, f.page, f.shift])).toEqual([
    [0, 8, 0, -64],
    [8, 8, 1, 0],
  ]);
  expect(plan.pages).toHaveLength(2);
});
it('counts automatic parity blanks toward the physical page budget', () => {
  const f = fixture([499 * 8, 1]);
  f.input.sections.push({ ...f.input.sections[0], id: 's2', start: 'oddPage' });
  f.input.blocks[1].section = 1;
  expect(planWordFragments(f.input, f.metrics)).toBeNull();
  f.input.sections[1].start = 'evenPage';
  expect(planWordFragments(f.input, f.metrics)!.pages).toHaveLength(500);
});
it('rebases all preceding fragments without displacing the following paragraph', () => {
  const f = fixture([17, 3]);
  f.metrics[0].lines[16].to = 16;
  f.input.blocks[0].to--;
  f.input.blocks[1].from--;
  f.input.blocks[1].to--;
  const plan = planWordFragments(f.input, f.metrics)!;
  expect(plan.blocks.map((b) => b.top)).toEqual([148, 488]);
  expect(plan.fragments.map((f) => [f.block, f.from, f.to, f.page, f.shift])).toEqual([
    [0, 0, 8, 0, -128],
    [0, 8, 16, 1, -64],
    [0, 16, 16, 2, 0],
    [1, 0, 3, 2, 0],
  ]);
});
