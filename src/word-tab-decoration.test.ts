import { expect, it } from 'vitest';
import reference from '../tests/fixtures/word-pdf-decorations/reference.json';
import { wordTabDecoration } from './word-tab-decoration';

it('uses the independently observed font decoration geometry for all 256 face/size controls', () => {
  expect(reference.expected).toHaveLength(256);
  for (const row of reference.expected) {
    const paint = wordTabDecoration({ family: row.family, size: row.size,
      weight: row.face === 'bold' || row.face === 'boldItalic' ? '700' : '400',
      style: row.face === 'italic' || row.face === 'boldItalic' ? 'italic' : 'normal',
      stretch: '100%', letterSpacing: 'normal', normalRatio: row.family === 'Arial' ? 1.1499 : 1.2207,
    }, row.kind === 'underline' ? 1 : 2)!;
    expect(paint).not.toBeNull();
    expect(Math.abs(paint.top * .75 + row.baseline - row.observed.top)).toBeLessThanOrEqual(.15);
    expect(Math.abs(paint.height * .75 - (row.observed.bottom - row.observed.top))).toBeLessThanOrEqual(.15);
    expect(Number.isInteger(paint.paintHeight)).toBe(true);
    expect(Math.abs(paint.paintHeight * .16 - paint.height)).toBeLessThan(1e-10);
    expect(paint.background).toMatch(/^linear-gradient\(to bottom, currentColor 0px, currentColor /);
  }
});

it('keeps unknown tab fonts and malformed mark profiles out of the native geometry path', () => {
  const font = { family: 'Arial', size: 10, weight: '400', style: 'normal', stretch: '100%',
    letterSpacing: 'normal', normalRatio: 1.1499 };
  for (const patch of [{ family: 'Inter' }, { normalRatio: NaN }, { normalRatio: 1.2 },
    { size: 0 }, { size: 401 }, { weight: '300' }, { style: 'oblique' }, { letterSpacing: '1px' },
    { script: 'superscript' as const }]) expect(wordTabDecoration({ ...font, ...patch }, 3)).toBeNull();
  for (const mark of [0, 4, 1.5, NaN]) expect(wordTabDecoration(font, mark)).toBeNull();
  const both = wordTabDecoration(font, 3)!;
  expect(both.rules).toHaveLength(2);
  expect(both.background).toContain('transparent');
  expect(both.top).toBeLessThan(0);
  expect(both.top + both.height).toBeGreaterThan(0);
});
