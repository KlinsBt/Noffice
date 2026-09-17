import { describe, expect, it } from 'vitest';
import reference from '../tests/fixtures/native-word-line-spacing-input.json';
import { wordLineSpacing, wordLineSpacingInput, type WordLineRule } from './word-line-spacing';

describe('independently measured native Word numeric line spacing', () => {
  it.each(reference.rows)('$receipt $rule $amount', row => {
    expect(wordLineSpacingInput(row.amount, row.rule as WordLineRule)).toEqual(row.expected);
  });

  it('preserves imported values below the authoring minimum while rejecting their numeric reapplication', () => {
    for (const row of reference.imports) {
      const { rule, line } = row.sample;
      const value = rule === 'auto' ? String(line / 240) : `${line / 20}pt`;
      expect(wordLineSpacing(value)).toEqual({ rule, line });
      expect(wordLineSpacingInput(line / (rule === 'auto' ? 240 : 20), rule as WordLineRule)).toBeNull();
      expect(row.reapplyAccepted).toBe(false);
    }
  });

  it('rejects nonnumeric and nonfinite input without coercion', () => {
    for (const value of [undefined, null, '', '1', {}, NaN, Infinity, -Infinity, -1, 0])
      expect(wordLineSpacingInput(value, 'auto')).toBeNull();
  });
});
