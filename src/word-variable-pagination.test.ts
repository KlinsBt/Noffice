import { it, expect } from 'vitest';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import native from '../tests/fixtures/native-word-variable-pagination.json';
import { planWordFragments, type ParagraphFlow } from './word-fragment-plan';
import { wordUniformLineHeight } from './word-font-line-metrics';
import type { SurfaceInput } from './word-section-surfaces';

for (const name of ['mixed', 'single', 'keep'] as const)
  it(`allocates every native character through variable heights and keep rules: ${name}`, () => {
    const reference = native.cases[name];
    expect(
      createHash('sha256')
        .update(fs.readFileSync(`tests/fixtures/word-variable-${name}.docx`))
        .digest('hex'),
    ).toBe(reference.sourceSha256);
    const section = (id: string) => ({
      id,
      width: 6000,
      height: 4200,
      margins: { top: 600, bottom: 600, left: 450, right: 450, header: 0, footer: 0, gutter: 0 },
      orientation: 'portrait' as const,
      paragraphs: [],
      headers: { default: null, even: null, first: null },
      footers: { default: null, even: null, first: null },
      differentFirstPage: false,
      simpleBody: true,
      start: 'nextPage',
    });
    const sections = name === 'single' ? [section('s0')] : [section('s0'), section('s1')];
    let from = 0;
    const input: SurfaceInput = {
      sections,
      width: 400,
      blocks: reference.snapshot.paragraphs.map((p, i) => {
        const block = {
          from,
          to: from + p.text.length + 1,
          width: 340,
          section: name === 'single' ? 0 : i < 6 ? 0 : 1,
          keepNext: name === 'keep' && [1, 9].includes(i),
          pageBreakBefore: name === 'keep' && i === 3,
        };
        from = block.to;
        return block;
      }),
    };
    const measurements: ParagraphFlow[] = reference.snapshot.paragraphs.map((p, i) => {
      const spacing = {
        rule: i < 9 ? ('auto' as const) : ('atLeast' as const),
        line: [240, 360, 480, 1000][Math.floor(i / 3)],
      };
      const sizes = i % 3 === 2 ? [10, 10] : [30, 20];
      let top = 0;
      const lines = sizes.map((size, index) => {
        const height = wordUniformLineHeight(size / 0.75, 2355 / 2048, spacing);
        const line = {
          from: index === 0 ? 0 : p.text.indexOf('\v') + 1,
          to: index === 0 ? p.text.indexOf('\v') + 1 : p.text.length - 1,
          top,
          height,
        };
        top += height;
        return line;
      });
      return {
        height: top,
        before: 0,
        after: 10 / 0.75,
        keepLines: !!p.keep,
        widowControl: !!p.widow,
        lines,
      };
    });
    const plan = planWordFragments(input, measurements)!;
    expect(plan).not.toBeNull();
    expect(plan.pages).toHaveLength(reference.snapshot.pages);
    for (const [index, p] of reference.snapshot.paragraphs.entries())
      for (const line of p.lines)
        for (let offset = line.from; offset < line.to; offset++) {
          const fragments = plan.fragments.filter(
            (f) => f.block === index && offset >= f.from && offset < f.to,
          );
          expect(fragments, `${index}:${offset}`).toHaveLength(1);
          expect(fragments[0].page + 1, `${index}:${offset}`).toBe(line.page);
        }
  });
