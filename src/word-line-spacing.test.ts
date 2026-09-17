import { expect, it } from 'vitest';
import { wordLineSpacing } from './word-line-spacing';
import { wordJSON } from './word-extensions';
import { sanitizeHTML } from './formats';

it.each([
  ['12pt', undefined, 'exact', 240],
  ['24px', undefined, 'exact', 360],
  ['18pt', 'atLeast', 'atLeast', 360],
  ['1.5', 'atLeast', 'auto', 360],
  ['1.15', undefined, 'auto', 276],
])(
  'normalizes %s with %s without confusing physical spacing and multiples',
  (value, rule, expected, line) => {
    expect(wordLineSpacing(value, rule)).toEqual({ rule: expected, line });
  },
);
it('rejects malformed, nonfinite, nonpositive and out-of-range values', () => {
  for (const value of [
    '',
    'normal',
    'calc(2px)',
    'NaN',
    'Infinity',
    '0',
    '-1pt',
    '2000pt',
    '0.00001pt',
    12,
  ])
    expect(wordLineSpacing(value)).toBeNull();
});
it('keeps minimum spacing through sanitization and ignores markers on automatic spacing', () => {
  const html = sanitizeHTML(
    '<p data-word-line-rule="atLeast" style="line-height:18pt">Minimum</p><p data-word-line-rule="atLeast" style="line-height:1.5">Multiple</p><p style="line-height:12pt">Legacy exact</p>',
  );
  expect(
    wordJSON(html).content!.map((p) => [p.attrs!.paragraphLineHeight, p.attrs!.paragraphLineRule]),
  ).toEqual([
    ['18pt', 'atLeast'],
    ['1.5', null],
    ['12pt', null],
  ]);
});
