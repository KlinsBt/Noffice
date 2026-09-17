import { expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { wordInterMetrics } from './word-inter-metrics';
import { wordDefaultFontBytes } from './word-default-font';
import type { WordLineSpacing } from './word-line-spacing';

const native = JSON.parse(
  readFileSync('tests/fixtures/native-word-inter-metrics.json', 'utf8'),
) as {
  fontHash: string;
  cases: { name: string; size: number; rule: string; starts: { page: number; y: number }[] }[];
};
const longNative = JSON.parse(
  readFileSync('tests/fixtures/native-word-inter-long-metrics.json', 'utf8'),
) as {
  fontHash: string;
  cases: {
    name: string;
    size: number;
    rule: string;
    starts: { page: number; line: number; y: number }[];
  }[];
};
const style = { fontFamily: 'Inter', fontWeight: '400', fontStyle: 'normal', fontStretch: '100%' };
const rules: Record<string, WordLineSpacing> = {
  single: { rule: 'auto', line: 240 },
  onehalf: { rule: 'auto', line: 360 },
  default: { rule: 'auto', line: 396 },
  double: { rule: 'auto', line: 480 },
  minimum: { rule: 'atLeast', line: 360 },
  exact: { rule: 'exact', line: 360 },
};

it('binds the native profile to the bundled font bytes actually exported', async () => {
  expect(
    createHash('sha256')
      .update(await wordDefaultFontBytes())
      .digest('hex'),
  ).toBe(native.fontHash);
  expect(longNative.fontHash).toBe(native.fontHash);
});

for (const sample of native.cases)
  it(`matches all native Inter line baselines at the original bound: ${sample.name}`, () => {
    const profile = wordInterMetrics(sample.size / 0.75, 2478 / 2048, style)!;
    const spacing = rules[sample.rule];
    expect(profile.supports(spacing)).toBe(true);
    const advance = profile.height(spacing);
    const before = spacing.rule === 'atLeast' ? Math.max(0, advance - profile.natural) : 0;
    expect(sample.starts).toHaveLength(12);
    for (const [i, line] of sample.starts.entries()) {
      expect(line.page).toBe(1);
      const baseline = 72 + (before + profile.ascent(spacing) + i * advance) * 0.75;
      expect(Math.abs(line.y - baseline)).toBeLessThanOrEqual(0.15);
    }
  });

for (const sample of longNative.cases)
  it(`retains native baselines across complete Inter pages: ${sample.name}`, () => {
    const profile = wordInterMetrics(sample.size / 0.75, 2478 / 2048, style)!;
    const spacing = rules[sample.rule];
    const advance = profile.height(spacing);
    const before = spacing.rule === 'atLeast' ? Math.max(0, advance - profile.natural) : 0;
    expect(sample.starts).toHaveLength(80);
    expect(sample.starts.at(-1)!.page).toBeGreaterThan(1);
    for (const line of sample.starts) {
      const baseline = 72 + (before + profile.ascent(spacing) + line.line * advance) * 0.75;
      expect(Math.abs(line.y - baseline)).toBeLessThanOrEqual(0.15);
    }
  });

it('retains fallback for unmeasured faces, sizes, styles, metrics and spacing rules', () => {
  for (const other of [{ fontFamily: 'Arial' }, { fontWeight: '700' }, { fontStyle: 'italic' }])
    expect(wordInterMetrics(16, 2478 / 2048, { ...style, ...other })).toBeNull();
  expect(wordInterMetrics(20, 2478 / 2048, style)).toBeNull();
  expect(wordInterMetrics(16, 1.1, style)).toBeNull();
  expect(wordInterMetrics(16, 2478 / 2048, style)!.supports({ rule: 'auto', line: 276 })).toBe(
    false,
  );
});
