import { expect, it } from 'vitest';
import { wordBodyGridOrigin } from './word-body-grid';
import reference from '../tests/fixtures/native-word-body-spacing-grid.json';

it('matches all60 independently shaded native body baselines including the before10/after6 failure', () => {
  for (const row of reference.rows) {
    let top = 288, previous = 0;
    row.offsets.forEach((glyph, i) => {
      const before = i === 2 ? row.before : 0, after = i === 2 ? row.after : 0;
      top += Math.max(previous, before);
      expect(Math.abs(wordBodyGridOrigin(top / .75, before / .75, after / .75) * .75 + 32.05 - glyph.y), `${row.name}/${glyph.text}`).toBeLessThanOrEqual(.15);
      top += 40; previous = after;
    });
  }
});

it('leaves nonfinite, negative and unmeasured spacing untouched', () => {
  for (const [before, after] of [[-1, 0], [0, -1], [NaN, 0], [0, Infinity], [54, 0], [0, 54]])
    expect(wordBodyGridOrigin(100, before, after)).toBe(100);
  expect(wordBodyGridOrigin(1, 2, 0)).toBe(1);
});
