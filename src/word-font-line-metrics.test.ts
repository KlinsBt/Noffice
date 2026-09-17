import { expect, it } from 'vitest';
import {
  wordUniformLineHeight,
  wordMixedLineMetrics,
  wordFontSizePixels,
  wordLineDisplayHeight,
  wordTextRectsShareLine,
  wordInlinePaintOrigin,
} from './word-font-line-metrics';

it('keeps fractional-width story paint origins stable across serialized transform round trips', () => {
  for (const scale of [0.5, 0.75, 1, 1.5, 2]) {
    expect(wordInlinePaintOrigin(0.00000025 / scale, scale)).toBe(0);
    expect(wordInlinePaintOrigin(-0.00000025 / scale, scale)).toBe(0);
    expect(wordInlinePaintOrigin(1.00000025 / scale, scale)).toBe(0);
    expect(wordInlinePaintOrigin((3 + 1 / 64) / scale, scale)).toBe(63 / 64 / scale);
    expect(wordInlinePaintOrigin(-1 / 64 / scale, scale)).toBe(1 / 64 / scale);
  }
});

it('distinguishes overlapping Calibri soft rows from multiple rectangles on one row', () => {
  expect(
    wordTextRectsShareLine([
      { top: 0, width: 200 },
      { top: 17.9375, width: 150 },
    ]),
  ).toBe(false);
  expect(
    wordTextRectsShareLine([
      { top: 0, width: 2 },
      { top: 0, width: 150 },
    ]),
  ).toBe(true);
  expect(
    wordTextRectsShareLine([
      { top: 0, width: 2 },
      { top: 8.96875, width: 75 },
    ]),
  ).toBe(false);
  expect(
    wordTextRectsShareLine([
      { top: 0, width: 100 },
      { top: 18, width: 0 },
    ]),
  ).toBe(true);
  expect(wordTextRectsShareLine([])).toBe(false);
});

it('retains cumulative native advances for long fractional hard lines at all supported zooms', () => {
  const nativeHeight = 11.5 / 0.75;
  for (const scale of [0.5, 0.6, 0.7, 0.8, 0.9, 1, 1.1, 1.2, 1.3, 1.4, 1.5]) {
    let physical = 0;
    for (let line = 0; line < 1000; line++) {
      physical += wordLineDisplayHeight(line * nativeHeight, nativeHeight, scale);
      expect(Math.abs(physical - (line + 1) * nativeHeight)).toBeLessThanOrEqual(
        1 / (scale * 64) + 1e-8,
      );
    }
  }
  // The previous per-line flooring drifts outside the measurement sanity bound
  // in the actual 26-line, 50% native fixture, disabling otherwise valid pages.
  expect(Math.abs((26 * Math.floor(nativeHeight * 32)) / 32 - 26 * nativeHeight)).toBeGreaterThan(
    0.2,
  );
});

it('matches native Arial minimum and automatic advances from measured font metrics', () => {
  const ratio = 2355 / 2048; // Installed Arial GDI source measurement, not a runtime constant.
  expect(wordUniformLineHeight(40, ratio, { rule: 'atLeast', line: 360 }) * 0.75).toBe(34.5);
  expect(wordUniformLineHeight(40 / 3, ratio, { rule: 'auto', line: 360 }) * 0.75).toBe(17.25);
  expect(wordUniformLineHeight(40, ratio, { rule: 'exact', line: 360 }) * 0.75).toBe(18);
  expect(wordUniformLineHeight(40, ratio, { rule: 'atLeast', line: 1000 }) * 0.75).toBe(50);
});

it('reads semantic font units without using the zero-sized view strut', () => {
  expect(wordFontSizePixels('40pt')).toBeCloseTo(160 / 3);
  expect(wordFontSizePixels('16px')).toBe(16);
  for (const value of ['0pt', '-1pt', 'NaNpx', '1em', 'calc(2px)', null])
    expect(wordFontSizePixels(value)).toBeNull();
});

it('keeps automatic extra leading below mixed text and minimum extra leading above it', () => {
  const ratio = 2355 / 2048;
  const sizes = [40 / 3, 40, 40 / 3];
  const auto = wordMixedLineMetrics(sizes, ratio, { rule: 'auto', line: 360 })!;
  expect(Math.max(...auto.heights) * 0.75).toBe(34.5);
  expect(auto.before).toBe(0);
  expect(auto.after * 0.75).toBe(17.25);
  const minimum = wordMixedLineMetrics(sizes, ratio, { rule: 'atLeast', line: 1000 })!;
  expect((Math.max(...minimum.heights) + minimum.before) * 0.75).toBe(50);
  expect(minimum.after).toBe(0);
  expect(wordMixedLineMetrics(sizes, ratio, { rule: 'exact', line: 360 })).toBeNull();
  expect(wordMixedLineMetrics(sizes, ratio, { rule: 'auto', line: 120 })).toBeNull();
  expect(wordMixedLineMetrics([NaN], ratio, { rule: 'auto', line: 240 })).toBeNull();
});
