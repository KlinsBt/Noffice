import { describe, expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import oracle from '../tests/fixtures/number-formats-oracle.json';
import edges from '../tests/fixtures/number-format-edges-oracle.json';
import { formatNumber } from './number-format';

describe('native Excel number-format oracle', () => {
  it('records the exact authored fixture', () => {
    expect(
      createHash('sha256')
        .update(readFileSync('tests/fixtures/number-formats-input.json'))
        .digest('hex'),
    ).toBe(oracle.inputSha256);
  });
  it.each(oracle.cases)('$code ($input)', ({ code, input, date1904, value, kind }) => {
    if (kind === 'error') expect(() => formatNumber(code, input, date1904)).toThrow(String(value));
    else expect(formatNumber(code, input, date1904)).toBe(value);
  });
});

describe('additional native Excel format boundaries', () => {
  it('records the exact edge fixture', () => {
    expect(
      createHash('sha256')
        .update(readFileSync('tests/fixtures/number-format-edges-input.json'))
        .digest('hex'),
    ).toBe(edges.inputSha256);
  });
  it.each(edges.cases)('$code ($input)', ({ code, input, value }) => {
    expect(formatNumber(code, input)).toBe(value);
  });
});

it('retains text literals and text sections without injecting markup', () => {
  expect(formatNumber('0;0;0;"text: "@', '<img>')).toBe('text: <img>');
  expect(formatNumber('"a/p s.00 "0.0', 1.5)).toBe('a/p s.00 1.5');
});
it('bounds code length, decimal precision and fraction search', () => {
  expect(() => formatNumber('0'.repeat(2049), 1)).toThrow('#LIMIT!');
  expect(() => formatNumber('0.' + '0'.repeat(101), 1)).toThrow('#LIMIT!');
  expect(() => formatNumber('# ?/?????', 0.314159)).toThrow('#LIMIT!');
});
