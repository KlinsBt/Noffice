import { expect, it } from 'vitest';
import { reflowWordColumns } from './word-column-reflow';
import type { ParagraphFlow } from './word-fragment-plan';
import type { SurfaceInput } from './word-section-surfaces';

const input = {
  width: 160,
  sections: [
    {
      id: 'section',
      width: 2400,
      height: 600,
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      columns: { widths: [900, 1500], spaces: [0, 0] },
    },
  ],
  blocks: [{ from: 0, to: 32, section: 0, width: 60 }],
} as SurfaceInput;
const initial = (): ParagraphFlow[] => [
  {
    height: 200,
    before: 0,
    after: 0,
    keepLines: false,
    widowControl: false,
    lines: Array.from({ length: 10 }, (_, i) => ({
      from: i * 3,
      to: i * 3 + 3,
      top: i * 20,
      height: 20,
    })),
  },
];

it('remeasures at new offsets that were not boundaries in the narrow-column layout', () => {
  const flows = initial(),
    before = JSON.stringify(flows);
  const result = reflowWordColumns(
    input,
    flows,
    new Map([
      [
        0,
        (from, width) => ({
          from,
          to: Math.min(30, from + width / 20),
          height: 20,
        }),
      ],
    ]),
  );
  expect(result).not.toBeNull();
  expect(
    result!.plan.fragments.map(({ from, to, page, column }) => ({ from, to, page, column })),
  ).toEqual([
    { from: 0, to: 6, page: 0, column: 0 },
    { from: 6, to: 16, page: 0, column: 1 },
    { from: 16, to: 22, page: 1, column: 0 },
    { from: 22, to: 30, page: 1, column: 1 },
  ]);
  expect(result!.flows[0].lines.some((line) => line.from === 11)).toBe(true);
  expect(JSON.stringify(flows)).toBe(before);
});

it('rejects an unstable measurement cycle without mutating the previous layout', () => {
  const flows = initial(),
    before = JSON.stringify(flows);
  let count = 3;
  const result = reflowWordColumns(
    input,
    flows,
    new Map([
      [
        0,
        (from) => {
          if (from === 0) count = count === 3 ? 5 : 3;
          return { from, to: Math.min(30, from + count), height: 20 };
        },
      ],
    ]),
  );
  expect(result).toBeNull();
  expect(JSON.stringify(flows)).toBe(before);
});
