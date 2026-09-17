import { describe, expect, it } from 'vitest';
import native from '../tests/fixtures/native-word-justification-arial-ten.json';
import calibri from '../tests/fixtures/native-word-justification-calibri-lines.json';
import { wordJustifiedLines } from './word-justified-lines';

describe('native Arial10 regular/bold modern line choices', () => {
  it.each(native.cases)('$name', (row) => {
    const advances = native.fonts.find((font) => font.face === row.face)!.advances;
    const measure = (text: string) => [...text].reduce((sum, c) =>
      sum + (advances as Record<string, number>)[c], 0);
    const lines = wordJustifiedLines(row.text, row.width, measure)!;
    expect(lines.map((line) => line.textEnd)).toEqual(row.ends);
    expect(lines.map((line) => row.text.slice(line.from, line.to)).join('')).toBe(row.text);
    expect(lines.at(-1)!.wordSpacing).toBe(0);
  });
});

describe('independent native Calibri11 regular ligature line choices', () => {
  it.each(calibri.cases)('$name', (row) => {
    const lines = wordJustifiedLines(row.text, row.width,
      text => (calibri.advances as Record<string, number>)[text])!;
    expect(lines.map(line => line.textEnd)).toEqual(row.ends);
    expect(lines.map(line => row.text.slice(line.from, line.to)).join('')).toBe(row.text);
    expect(lines.at(-1)!.wordSpacing).toBe(0);
  });
});

it.each(['', ' lead', 'tail ', 'two  spaces', 'a\tb', 'a\nb', 'caf\u00e9', 'a'.repeat(20001)])(
  'retains the fallback for unqualified text %j', (text) => {
    expect(wordJustifiedLines(text, 100, (part) => part.length)).toBeNull();
  },
);
it.each([0, -1, NaN, Infinity])('rejects invalid available width %s', (width) => {
  expect(wordJustifiedLines('small words', width, (part) => part.length)).toBeNull();
});
it.each([0, -1, NaN, Infinity])('rejects invalid logical advances %s', (advance) => {
  expect(wordJustifiedLines('small words', 100, () => advance)).toBeNull();
});
it('does not guess character splitting for an oversized word', () => {
  expect(wordJustifiedLines('unbroken tiny', 5, (part) => part.length)).toBeNull();
});
