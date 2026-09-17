import { it, expect } from 'vitest';
import { wordNativeBaseline } from './word-native-baseline';
import { wordUniformLineHeight } from './word-font-line-metrics';
import native from '../tests/fixtures/native-word-glyph-fonts.json';
import exact from '../tests/fixtures/native-word-exact-baselines.json';
import authored from '../tests/fixtures/native-word-authored-pagination.json';
import calibri from '../tests/fixtures/native-word-story-calibri-metrics.json';
import leaders from '../tests/fixtures/native-word-tab-leader-baselines.json';
import storyRanges from '../tests/fixtures/native-word-story-font-range-geometry.json';
import fontSteps from '../tests/fixtures/native-word-story-font-step-geometry.json';
import decorations from '../tests/fixtures/word-pdf-decorations/baselines.json';
import hyphens from '../tests/fixtures/word-soft-hyphens/baselines.json';
import sectionEnds from '../tests/fixtures/word-empty-section-end/reference.json';
import firstSection from '../tests/fixtures/word-first-section-parity/reference.json';
import listBaselines from '../tests/fixtures/word-section-containers/baselines.json';

it('matches all independently shaded Arial12 Single list and paragraph origins', () => {
  const baseline=wordNativeBaseline(16,18.4,1.1499,'Arial','400','normal','normal',{rule:'auto',line:240});
  expect(baseline).not.toBeNull();
  expect(listBaselines.rows).toHaveLength(5);
  for(const page of listBaselines.rows) {
    expect(page.unchangedGlyphs).toBe(true);
    for(const sample of page.offsets) expect(Math.abs(baseline!*.75-sample.offset)).toBeLessThan(.15);
  }
  expect(wordNativeBaseline(16,18.4,1.1499,'Arial','700','normal','normal',{rule:'auto',line:240})).toBeNull();
  expect(wordNativeBaseline(16,18.4,1.1499,'Arial','400','italic','normal',{rule:'auto',line:240})).toBeNull();
});

it('matches the native regular Arial12 first-paragraph baseline at 1.65 spacing', () => {
  const control = firstSection.baselineControl;
  const args = [
    16,
    16 * 1.1499,
    1.1499,
    'Arial',
    '400',
    'normal',
    'normal',
    { rule: 'auto', line: 396 },
  ] as const;
  const baseline = wordNativeBaseline(...args);
  expect(baseline).not.toBeNull();
  for (const sample of control.samples)
    expect(Math.abs(baseline! * 0.75 - sample.offset)).toBeLessThan(0.15);
  expect(
    wordNativeBaseline(16, 16 * 1.1499, 1.1499, 'Arial', '700', 'normal', 'normal', args[7]),
  ).toBeNull();
  expect(
    wordNativeBaseline(16, 16 * 1.1499, 1.1499, 'Arial', '400', 'normal', 'condensed', args[7]),
  ).toBeNull();
});

it.each(sectionEnds.baselineControls)('matches empty-section typing baseline: $name', (row) => {
  const args = [
    row.size / 0.75,
    (row.size * 1.1499) / 0.75,
    1.1499,
    'Arial',
    '400',
    'normal',
    'normal',
    { rule: 'exact', line: row.line },
  ] as const;
  expect(Math.abs(wordNativeBaseline(...args)! * 0.75 - row.offset)).toBeLessThanOrEqual(0.15);
  expect(
    wordNativeBaseline(args[0], args[1], args[2], 'Arial', '700', 'normal', 'normal', args[7]),
  ).toBeNull();
  expect(
    wordNativeBaseline(args[0], args[1], args[2], 'Arial', '400', 'normal', 'condensed', args[7]),
  ).toBeNull();
});

it('matches245 independent exact24 line origins, including both hyphen encodings and wrapped pages', () => {
  expect(hyphens.rows).toHaveLength(245);
  expect(hyphens.pages).toHaveLength(18);
  expect(hyphens.pages.every((page) => page.glyphsIdentical)).toBe(true);
  for (const row of hyphens.rows) {
    const ratio = row.family === 'Calibri' ? 1.2207 : 1.1499;
    const baseline = wordNativeBaseline(
      row.size / 0.75,
      (row.size * ratio) / 0.75,
      ratio,
      row.family,
      ['bold', 'boldItalic'].includes(row.face) ? '700' : '400',
      ['italic', 'boldItalic'].includes(row.face) ? 'italic' : 'normal',
      'normal',
      { rule: 'exact', line: 480 },
    );
    expect(baseline).not.toBeNull();
    expect(Math.abs(baseline! * 0.75 - row.offset)).toBeLessThanOrEqual(0.15);
  }
});

it('matches74 independent shaded paragraph origins across all eight decoration font faces', () => {
  expect(decorations.rows).toHaveLength(74);
  expect(decorations.pages.every((page) => page.glyphsIdentical)).toBe(true);
  for (const row of decorations.rows) {
    const ratio = row.family === 'Calibri' ? 1.2207 : 1.1499;
    const baseline = wordNativeBaseline(
      row.size / 0.75,
      (row.size * ratio) / 0.75,
      ratio,
      row.family,
      ['bold', 'boldItalic'].includes(row.face) ? '700' : '400',
      ['italic', 'boldItalic'].includes(row.face) ? 'italic' : 'normal',
      'normal',
      { rule: 'exact', line: 800 },
    );
    expect(baseline).not.toBeNull();
    expect(Math.abs(baseline! * 0.75 - row.offset)).toBeLessThanOrEqual(0.15);
  }
});

it('matches independently shaded Arial9/11 exact40 Grow/Shrink paragraphs', () => {
  for (const row of fontSteps.rows) {
    const size = row.name.includes('Grow') ? 11 : 9;
    const baseline = wordNativeBaseline(
      size / 0.75,
      (size * 1.1499) / 0.75,
      1.1499,
      'Arial',
      '400',
      'normal',
      'normal',
      { rule: 'exact', line: 800 },
    );
    expect(baseline).not.toBeNull();
    for (const page of row.pages) {
      expect(page.glyphsUnchanged).toBe(true);
      for (const offset of page.offsets)
        expect(Math.abs(baseline! * 0.75 - offset.offset)).toBeLessThanOrEqual(0.15);
    }
  }
});

it('retains the independently shaded mixed Arial10/Times20 exact40 story baseline', () => {
  for (const row of storyRanges.rows)
    for (const page of row.pages) {
      expect(page.glyphsUnchanged).toBe(true);
      for (const [family, size] of [
        ['Arial', 10],
        ['Times New Roman', 20],
      ] as const) {
        const baseline = wordNativeBaseline(
          size / 0.75,
          (size * 1.1499) / 0.75,
          1.1499,
          family,
          '400',
          'normal',
          'normal',
          { rule: 'exact', line: 800 },
        );
        expect(baseline).not.toBeNull();
        for (const offset of page.offsets)
          expect(Math.abs(baseline! * 0.75 - offset.offset)).toBeLessThanOrEqual(0.15);
      }
    }
});

it.each(leaders.rows)('matches the independently shaded exact40 baseline: $name', (row) => {
  const ratio = row.font === 'Calibri' ? 1.2207 : 1.1499;
  for (const size of new Set([row.size, row.tabSize || row.size])) {
    const baseline = wordNativeBaseline(
      size / 0.75,
      (size * ratio) / 0.75,
      ratio,
      row.font,
      row.bold ? '700' : '400',
      row.italic ? 'italic' : 'normal',
      'normal',
      { rule: 'exact', line: 800 },
    );
    expect(baseline).not.toBeNull();
    expect(Math.abs(baseline! * 0.75 - row.offset)).toBeLessThanOrEqual(0.15);
  }
});

it('matches native Calibri11 header/footer baselines across all sixteen authored contexts', () => {
  expect(calibri.rows).toHaveLength(16);
  for (const row of calibri.rows) {
    const spacing =
      row.mode === 'double'
        ? { rule: 'auto' as const, line: 480 }
        : row.mode === 'minimum30'
          ? { rule: 'atLeast' as const, line: 600 }
          : row.mode === 'exact12'
            ? { rule: 'exact' as const, line: 240 }
            : { rule: 'auto' as const, line: 240 };
    const natural = wordUniformLineHeight(11 / 0.75, calibri.normalRatio, {
      rule: 'auto',
      line: 240,
    });
    const height = wordUniformLineHeight(11 / 0.75, calibri.normalRatio, spacing);
    const baseline = wordNativeBaseline(
      11 / 0.75,
      natural,
      calibri.normalRatio,
      'Calibri',
      '400',
      'normal',
      'normal',
      spacing,
    );
    expect(baseline).not.toBeNull();
    for (const page of row.paintPages) {
      if (!page.lines.length) continue;
      const origin = row.kind === 'Headers' ? 18 : 220 - 18 - page.lines.length * height * 0.75;
      const excess = spacing.rule === 'atLeast' ? Math.max(0, height - natural) : 0;
      for (const [index, line] of page.lines.entries())
        expect(
          Math.abs(origin + (index * height + baseline! + excess) * 0.75 - line.y),
          row.name,
        ).toBeLessThanOrEqual(0.15);
    }
  }
});

it('does not extrapolate the Calibri profile to unmeasured font descriptors or spacing', () => {
  const base = [
    11 / 0.75,
    13.4 / 0.75,
    calibri.normalRatio,
    'Calibri',
    '400',
    'normal',
    'normal',
    { rule: 'auto', line: 240 },
  ] as const;
  for (const [index, value] of [
    [0, NaN],
    [0, 12 / 0.75],
    [2, 1.2],
    [4, '700'],
    [5, 'italic'],
    [6, 'condensed'],
    [7, { rule: 'auto', line: 360 }],
    [7, { rule: 'exact', line: 300 }],
    [7, { rule: 'atLeast', line: 1000 }],
  ] as const) {
    const args: any[] = [...base];
    args[index] = value;
    expect(wordNativeBaseline(...(args as Parameters<typeof wordNativeBaseline>))).toBeNull();
  }
});

it('matches the independently authored exact12 baseline across both pages and three edit states', () => {
  expect(authored.baselines).toHaveLength(210);
  const baseline = wordNativeBaseline(
    10 / 0.75,
    (10 / 0.75) * native.normalRatio,
    native.normalRatio,
    'Arial',
    '400',
    'normal',
    'normal',
    { rule: 'exact', line: 240 },
  );
  expect(baseline).not.toBeNull();
  for (const line of authored.baselines)
    expect(Math.abs(baseline! * 0.75 - line.baselineFromLineTop)).toBeLessThanOrEqual(0.15);
});

it('matches all twelve exact native baselines relative to independently measured paragraph boxes', () => {
  expect(exact.paragraphBoxes.glyphsUnchanged).toBe(true);
  for (const row of exact.baselines) {
    const baseline = wordNativeBaseline(
      row.size / 0.75,
      (row.size / 0.75) * native.normalRatio,
      native.normalRatio,
      'Arial',
      '400',
      'normal',
      'normal',
      { rule: 'exact', line: row.line * 20 },
    );
    expect(baseline).not.toBeNull();
    expect(Math.abs(baseline! * 0.75 - row.baselineFromLayoutTop)).toBeLessThanOrEqual(0.15);
  }
});

it('matches all 28 independently measured native Arial font/spacing baselines', () => {
  for (const row of native.rows) {
    const spacing = { rule: row.rule as 'auto' | 'atLeast', line: row.line };
    const natural = wordUniformLineHeight(row.size / 0.75, native.normalRatio, {
      rule: 'auto',
      line: 240,
    });
    const baseline = wordNativeBaseline(
      row.size / 0.75,
      natural,
      native.normalRatio,
      'Arial, Inter',
      '400',
      'normal',
      'normal',
      spacing,
    );
    expect(baseline).not.toBeNull();
    const before = spacing.rule === 'atLeast' ? Math.max(0, spacing.line / 15 - natural) : 0;
    expect(Math.abs((baseline! + before) * 0.75 - row.ascent)).toBeLessThan(0.15);
  }
});

it('retains fallback for unmeasured fonts, sizes, spacing and invalid metrics', () => {
  const base = [
    80 / 3,
    92 / 3,
    native.normalRatio,
    'Arial',
    '400',
    'normal',
    'normal',
    { rule: 'auto', line: 360 },
  ] as const;
  for (const [index, value] of [
    [0, 28],
    [0, NaN],
    [1, Infinity],
    [2, 1.2],
    [3, 'Inter'],
    [4, '700'],
    [5, 'italic'],
    [6, 'condensed'],
    [7, { rule: 'auto', line: 276 }],
    [7, { rule: 'atLeast', line: 1020 }],
    [7, { rule: 'exact', line: 1000 }],
  ] as const) {
    const args: any[] = [...base];
    args[index] = value;
    expect(wordNativeBaseline(...(args as Parameters<typeof wordNativeBaseline>))).toBeNull();
  }
});
