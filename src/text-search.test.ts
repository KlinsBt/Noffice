import { expect, it } from 'vitest';
import { textMatches, replaceMatches } from './text-search';

it('matches literal metacharacters and preserves replacement dollar strings', () => {
  const text = 'Price $5.00? Price $5.00?';
  const found = textMatches(text, '$5.00?');
  expect(found).toEqual([
    { from: 6, to: 12 },
    { from: 19, to: 25 },
  ]);
  expect(replaceMatches(text, found, '$&', 100)).toBe('Price $& Price $&');
});
it('keeps Unicode offsets correct and supports case and whole word options', () => {
  expect(textMatches('İ test TEST testing', 'test', { wholeWord: true })).toEqual([
    { from: 2, to: 6 },
    { from: 7, to: 11 },
  ]);
  expect(textMatches('test TEST', 'test', { matchCase: true })).toHaveLength(1);
  expect(textMatches('𐐀test test𐐀 test', 'test', { wholeWord: true })).toEqual([
    { from: 14, to: 18 },
  ]);
  expect(textMatches('café caféine', 'café', { wholeWord: true })).toHaveLength(1);
});
it('matches entire cells, nonoverlapping occurrences and empty searches', () => {
  expect(textMatches('abc', 'ab', { wholeCell: true })).toEqual([]);
  expect(textMatches('ABC', 'abc', { wholeCell: true })).toHaveLength(1);
  expect(textMatches('aaaaa', 'aa')).toEqual([
    { from: 0, to: 2 },
    { from: 2, to: 4 },
  ]);
  expect(textMatches('abc', '')).toEqual([]);
});
it('rejects excessive work and oversized replacement before mutation', () => {
  expect(() => textMatches('a'.repeat(10001), 'a')).toThrow('10,000');
  expect(() => textMatches('a', 'x'.repeat(257))).toThrow('256');
  expect(() => textMatches('a', 'a', {}, { characters: 8000000, matches: 0 })).toThrow('too large');
  expect(() => replaceMatches('abc', [{ from: 0, to: 3 }], 'abcd', 3)).toThrow('limit');
});
