import { expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import spec from '../tests/fixtures/financial-input.json';
import oracle from '../tests/fixtures/financial-oracle.json';
import { calculator } from './formulas';
import type { Sheet } from './model';
const sheet: Sheet = {
  id: 'finance',
  name: spec.sheet,
  rows: 200,
  cols: 6,
  cells: Object.fromEntries([
    ...Object.entries(spec.cells).map(([ref, value]) => [
      ref,
      {
        value: String(value),
        ...(typeof value === 'string' ? { dataType: 'text' } : {}),
      },
    ]),
    ...spec.cases.map(({ ref, formula }) => [ref, { value: formula }]),
  ]),
};
it('binds financial comparisons to the exact native Excel fixture and formulas', () => {
  expect(
    createHash('sha256').update(readFileSync('tests/fixtures/financial-input.json')).digest('hex'),
  ).toBe(oracle.inputSha256);
  expect(oracle.cases.map(({ ref, formula }) => ({ ref, formula }))).toEqual(spec.cases);
});
it.each(oracle.cases)('native Excel $ref $formula', ({ ref, value, kind }) => {
  const result = calculator([sheet]).result(sheet, ref);
  if (kind === 'error') expect(result).toEqual({ kind: 'error', value });
  else {
    expect(result.kind).toBe('value');
    expect(typeof result.value).toBe('number');
    expect(Math.abs(Number(result.value) - Number(value))).toBeLessThanOrEqual(
      1e-10 * Math.max(1, Math.abs(Number(value))),
    );
  }
});
it('rejects invalid arity and propagates unsupported financial inputs', () => {
  const data = {
    ...sheet,
    cells: {
      ...sheet.cells,
      Z1: { value: '=PMT(0.1,10)' },
      Z2: { value: '=PV(UNKNOWN(),10,100)' },
      Z3: { value: '=EFFECT(0.1,12,1)' },
    },
  };
  const calc = calculator([data]);
  expect(calc(data, 'Z1')).toBe('#VALUE!');
  expect(calc.result(data, 'Z2')).toEqual({ kind: 'unsupported', value: '#UNSUPPORTED!' });
  expect(calc(data, 'Z3')).toBe('#VALUE!');
});
