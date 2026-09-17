import { expect, it } from 'vitest';
import { textMatches, replaceMatches } from './text-search';
it.each([
  ['sat set sit', 's?t', ['sat', 'set', 'sit']],
  ['xx a1b a2b yy', 'a*b', ['a1b', 'a2b']],
  ['any text', '*', ['any text']],
  ['fy91? fy91x', 'fy91~?', ['fy91?']],
  ['cost* and ~', '~*', ['*']],
  ['cost* and ~', '~~', ['~']],
  ['a\nb', 'a*b', ['a\nb']],
  ['a😀b', 'a?b', ['a😀b']],
] as const)(
  'matches Excel-style pattern %s / %s without regex backtracking',
  (text, query, expected) => {
    expect(
      textMatches(text, query, { wildcards: true }).map((m) => text.slice(m.from, m.to)),
    ).toEqual(expected);
  },
);
it('supports whole cells, case, literal mode and literal replacement strings', () => {
  expect(textMatches('extra SAT', 's?t', { wildcards: true, wholeCell: true })).toEqual([]);
  expect(textMatches('SAT', 's?t', { wildcards: true, matchCase: true })).toEqual([]);
  expect(textMatches('SAT', 's?t')).toEqual([]);
  expect(
    replaceMatches('sat set', textMatches('sat set', 's?t', { wildcards: true }), '$&', 100),
  ).toBe('$& $&');
  expect(textMatches('', '*', { wildcards: true })).toEqual([]);
});
it('bounds adversarial patterns before executing them', () => {
  expect(() => textMatches('a'.repeat(30000), '*a'.repeat(100) + 'b', { wildcards: true })).toThrow(
    'too complex',
  );
});
