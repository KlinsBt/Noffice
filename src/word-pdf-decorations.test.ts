import { describe, expect, it } from 'vitest';
import reference from '../tests/fixtures/word-pdf-decorations/reference.json';
import { wordPdfDecorationMetrics, wordPdfDecorationRules } from './word-pdf-decorations';
import { validateWordPdfSnapshot, type WordPdfSnapshot, type WordPdfDecorationSpan } from './word-pdf-model';

const glyph = () => ({ text: 'A', x: 72, y: 80, size: 10, family: 'Arial',
  color: [0, 0, 0] as [number, number, number], run: 1, decoration: 3, width: 5 });
const metrics = { unitsPerEm: 2048, underlinePosition: -217, underlineThickness: 150,
  strikePosition: 530, strikeThickness: 102 };
describe('PDF underline and strikethrough', () => {
  it('matches all256 independent held-out Word PDF vertical observations at the unchanged .15pt bound', () => {
    expect(reference.expected).toHaveLength(256);
    for (const row of reference.expected) {
      const font = row.metrics;
      const page = { width: 595, height: 842, glyphs: [], decorations: [{ ...glyph(), size: row.size, y: row.baseline,
        decoration: row.kind === 'underline' ? 1 : 2 }] };
      const [rule] = wordPdfDecorationRules(page, new Map([['Arial', {
        unitsPerEm: font.em, underlinePosition: font.underlinePosition, underlineThickness: font.underlineThickness,
        strikePosition: font.strikePosition, strikeThickness: font.strikeSize,
      }]]));
      expect(Math.abs(rule.y - row.observed.top)).toBeLessThanOrEqual(.15);
      expect(Math.abs(rule.height - (row.observed.bottom - row.observed.top))).toBeLessThanOrEqual(.15);
    }
  });
  it('retains authored spaces and divides paint at gaps, colors, lines and unmarked text', () => {
    const a = glyph();
    const page = { width: 595, height: 842, glyphs: [], decorations: [a, { ...a, text: ' ', x: 77, width: 2 },
      { ...a, x: 79, color: [1, 0, 0] as [number, number, number] },
      { ...a, x: 89, run: 2 }, { ...a, x: 94, y: 100, run: 2 }] };
    const before = structuredClone(page);
    const rules = wordPdfDecorationRules(page, new Map([['Arial', metrics]]));
    expect(rules).toHaveLength(8);
    expect(rules[0]).toMatchObject({ x: 72, width: 7, color: [0, 0, 0] });
    expect(rules[1]).toMatchObject({ x: 79, width: 5, color: [1, 0, 0] });
    expect(page).toEqual(before);
    expect(() => wordPdfDecorationRules(page, new Map())).toThrow(/unavailable/);
  });
  it('rejects malformed decoration inputs before worker drawing', () => {
    for (const patch of [{ decoration: 0 }, { decoration: 4 }, { decoration: 1.5 },
      { width: undefined }, { width: NaN }, { width: -1 }, { width: 3169 }]) {
      const snapshot: WordPdfSnapshot = { content: '', pages: [{ width: 595, height: 842, glyphs: [],
        decorations: [{ ...glyph(), ...patch } as WordPdfDecorationSpan] }] };
      expect(() => validateWordPdfSnapshot(snapshot)).toThrow();
    }
    expect(() => wordPdfDecorationMetrics(new Uint8Array(12))).toThrow(/metrics/);
    const invalid = new Uint8Array(12); new DataView(invalid.buffer).setUint16(4, 256);
    expect(() => wordPdfDecorationMetrics(invalid)).toThrow(/metrics/);
  });
  it('paints tab-only fonts and contiguous text/tab spans without emitting a tab glyph', () => {
    const a = glyph();
    const page = { width: 595, height: 842, glyphs: [], decorations: [
      { ...a, width: 6.7 }, { ...a, x: 78.7, width: 65.3 }, { ...a, x: 144, width: 6.6 },
    ] };
    expect(validateWordPdfSnapshot({ content: 'A\tB', pages: [page] })).toEqual(['Arial']);
    expect(wordPdfDecorationRules(page, new Map([['Arial', metrics]])))
      .toEqual([expect.objectContaining({ x: 72, width: 78.6 }), expect.objectContaining({ x: 72, width: 78.6 })]);
    expect(page.glyphs).toEqual([]);
    for (const patch of [{ run: -1 }, { run: .5 }, { face: 'missing' }, { family: '' },
      { size: 401 }, { x: Infinity }, { color: [2, 0, 0] }])
      expect(() => validateWordPdfSnapshot({ content: '\t', pages: [{ ...page,
        decorations: [{ ...a, ...patch } as WordPdfDecorationSpan] }] })).toThrow();
  });
});
