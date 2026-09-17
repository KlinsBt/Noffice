import { it, expect } from 'vitest';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import spec from '../tests/fixtures/aggregates-input.json';
import oracle from '../tests/fixtures/aggregates-oracle.json';
import { calculator } from './formulas';
import type { Sheet } from './model';
const sheet: Sheet = {
  id: 'aggregates',
  name: spec.sheet,
  rows: 150,
  cols: 6,
  definedNames: spec.names,
  cells: Object.fromEntries([
    ...Object.entries(spec.cells).map(([ref, value]) => [
      ref,
      {
        value: String(value),
        ...(typeof value === 'string' && !value.startsWith('=') ? { dataType: 'text' } : {}),
      },
    ]),
    ...spec.cases.map(({ ref, formula }) => [ref, { value: formula }]),
  ]),
};
it('binds aggregate expectations to exact native Excel input and formulas', () => {
  expect(
    createHash('sha256').update(readFileSync('tests/fixtures/aggregates-input.json')).digest('hex'),
  ).toBe(oracle.inputSha256);
  expect(oracle.cases.map(({ ref, formula }) => ({ ref, formula }))).toEqual(spec.cases);
});
it.each(oracle.cases)('native Excel $ref $formula', ({ ref, value, kind }) => {
  const result = calculator([sheet]).result(sheet, ref);
  expect(result.kind).toBe(kind);
  if (typeof value === 'number') expect(result.value).toBeCloseTo(value, 12);
  else expect(result.value).toBe(value);
});
it('does not disguise unsupported calculations or circular dependencies inside COUNT/COUNTA', () => {
  const scoped = {
    ...sheet,
    cells: {
      A1: { value: '=FutureFunction(1)' },
      B1: { value: '=COUNT(A1)' },
      B2: { value: '=COUNTA(A1)' },
      B3: { value: '=COUNT(B3)' },
    },
  };
  const calc = calculator([scoped]);
  expect(calc(scoped, 'B1')).toBe('#UNSUPPORTED!');
  expect(calc(scoped, 'B2')).toBe('#UNSUPPORTED!');
  expect(calc(scoped, 'B3')).toBe('#CYCLE!');
});
