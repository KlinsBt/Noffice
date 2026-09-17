import { expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { wordInterMetrics } from './word-inter-metrics';
import { planWordFragments, type ParagraphFlow } from './word-fragment-plan';
import type { SurfaceInput } from './word-section-surfaces';
import type { WordLineSpacing } from './word-line-spacing';

const native = JSON.parse(
  readFileSync('tests/fixtures/native-word-inter-long-metrics.json', 'utf8'),
) as {
  cases: {
    name: string;
    size: number;
    rule: string;
    sourceHash: string;
    starts: { page: number; text: string }[];
  }[];
};
const spacing: Record<string, WordLineSpacing> = {
  single: { rule: 'auto', line: 240 },
  onehalf: { rule: 'auto', line: 360 },
  default: { rule: 'auto', line: 396 },
  double: { rule: 'auto', line: 480 },
  minimum: { rule: 'atLeast', line: 360 },
  exact: { rule: 'exact', line: 360 },
};
for (const sample of native.cases)
  it(`allocates every independently captured full-page Inter line: ${sample.name}`, () => {
    expect(
      createHash('sha256')
        .update(readFileSync(`tests/fixtures/word-inter-long-metrics/${sample.name}.docx`))
        .digest('hex'),
    ).toBe(sample.sourceHash);
    const rule = spacing[sample.rule];
    const metrics = wordInterMetrics(sample.size / 0.75, 2478 / 2048, {
      fontFamily: 'Inter',
      fontWeight: '400',
      fontStyle: 'normal',
      fontStretch: '100%',
    })!;
    const advance = metrics.height(rule);
    let from = 0;
    const lines = sample.starts.map((line, i) => {
      const start = from;
      from += line.text.length + (i + 1 < sample.starts.length ? 1 : 0);
      return {
        from: start,
        to: from,
        top: i * advance,
        height: advance,
        fitHeight: rule.rule === 'auto' ? metrics.natural : advance,
      };
    });
    const input: SurfaceInput = {
      width: 11906 / 15,
      sections: [
        {
          id: 's0',
          width: 11906,
          height: 16838,
          margins: {
            left: 1440,
            right: 1440,
            top: 1440,
            bottom: 1440,
            header: 708,
            footer: 708,
            gutter: 0,
          },
          orientation: 'portrait',
          paragraphs: [],
          headers: { default: null, even: null, first: null },
          footers: { default: null, even: null, first: null },
          differentFirstPage: false,
          simpleBody: true,
          start: 'nextPage',
        },
      ],
      blocks: [{ from: 0, to: from + 2, section: 0, width: (11906 - 2880) / 15 }],
    };
    const flow: ParagraphFlow = {
      height: advance * lines.length,
      before: 0,
      after: 0,
      lines,
      keepLines: false,
      widowControl: true,
    };
    const plan = planWordFragments(input, [flow]);
    expect(plan).not.toBeNull();
    expect(
      lines.map(
        (line) => plan!.fragments.find((f) => f.from <= line.from && f.to >= line.to)!.page + 1,
      ),
    ).toEqual(sample.starts.map((line) => line.page));
  });
