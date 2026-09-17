import { describe, expect, it } from 'vitest';
import { wordTabBarOffsets, wordTabBarWidth } from './word-tab-bars';
import reference from '../tests/fixtures/native-word-tab-bars.json';

describe('native bar stop ink across fonts, spacing, empty and hard lines', () => {
  for (const [index, row] of reference.rows.entries())
    it(`${row.name} line ${index}: paints the independently measured line box`, () => {
      const stops = JSON.stringify(row.positions.map((position) => ({ position, alignment: 'bar', leader: 'none' })));
      const offsets = wordTabBarOffsets(stops, row.indent / 15);
      expect(offsets).toHaveLength(row.strokes.length);
      row.strokes.forEach((stroke, i) => {
        expect(stroke.fill).toBe(0);
        expect(stroke.stroke).toBe(1);
        expect(stroke.colors[1]).toEqual([0, 0, 0, 255]);
        const x = reference.textMargin + row.indent / 20 + offsets[i] * .75;
        const xs = stroke.segments.map((segment) => segment[0]);
        const ys = stroke.segments.map((segment) => segment[1]);
        expect(Math.max(...xs) - Math.min(...xs)).toBe(0);
        expect(Math.abs(x - (xs[0] - stroke.strokeWidth / 2))).toBeLessThanOrEqual(.15);
        expect(Math.abs(x + wordTabBarWidth * .75 - (xs[0] + stroke.strokeWidth / 2))).toBeLessThanOrEqual(.15);
        expect(Math.abs(row.origin - Math.min(...ys))).toBeLessThanOrEqual(.15);
        expect(Math.abs(row.origin + row.height - Math.max(...ys))).toBeLessThanOrEqual(.15);
      });
    });
});
