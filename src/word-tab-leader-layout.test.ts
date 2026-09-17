import { describe, it, expect } from 'vitest';
import native from '../tests/fixtures/native-word-tab-leader-geometry.json';
import { wordTabLeaderOrigins, wordTabLeaderPrintBounds } from './word-tab-leader-layout';
import boundaries from '../tests/fixtures/native-word-story-leader-boundaries.json';

it.each(boundaries.rows)('matches the independently captured printer-cell boundary $name', row => {
  const bounds = wordTabLeaderPrintBounds(row.left, row.advance, row.width);
  expect(bounds).not.toBeNull();
  const origins = wordTabLeaderOrigins(...bounds!, row.pitch);
  expect(origins).toHaveLength(row.count);
  expect(Math.abs(origins![0] - row.first)).toBeLessThanOrEqual(.15);
});

it('rejects invalid printer geometry before choosing leader cells', () => {
  for (const input of [[NaN, 1, 2], [1, Infinity, 2], [1, 2, -1], [1, 2, NaN]])
    expect(wordTabLeaderPrintBounds(...input as [number, number, number])).toBeNull();
});

describe('native tab leader page grid', () => {
  for (const row of native.rows) it(row.name, () => {
    const actual = wordTabLeaderOrigins(row.start, row.end, row.pitch);
    expect(actual).not.toBeNull();expect(actual).toHaveLength(row.expected.length);
    actual!.forEach((x, i) => expect(Math.abs(x - row.expected[i])).toBeLessThanOrEqual(.15));
  });
  it('keeps complete cells at exact endpoints and rejects genuinely partial cells', () => {
    expect(wordTabLeaderOrigins(.3, .6, .1)).toEqual([.30000000000000004, .4, .5]);
    expect(wordTabLeaderOrigins(.3 + 1e-8, .6 - 1e-8, .1)).toEqual([.4]);
    expect(wordTabLeaderOrigins(1, 1.1, 2)).toEqual([]);
    expect(wordTabLeaderOrigins(1, 1, 2)).toEqual([]);
  });
  it('bounds allocation and rejects invalid or unrepresentable geometry', () => {
    expect(wordTabLeaderOrigins(0, 5000, 1)).toHaveLength(5000);
    for (const args of [[0, 5001, 1], [0, 1, 0], [0, 1, -1], [2, 1, 1], [NaN, 1, 1],
      [0, Infinity, 1], [0, 1, Number.MIN_VALUE]]) expect(wordTabLeaderOrigins(...args as [number, number, number])).toBeNull();
  });
});
