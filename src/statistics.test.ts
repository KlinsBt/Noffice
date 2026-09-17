import { expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import spec from '../tests/fixtures/statistics-input.json';
import oracle from '../tests/fixtures/statistics-oracle.json';
import { calculator } from './formulas';
import type { Sheet } from './model';
const sheet: Sheet = {
  id: 'statistics',
  name: spec.sheet,
  rows: 120,
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
it('binds statistical expectations to the exact native Excel input', () => {
  expect(
    createHash('sha256').update(readFileSync('tests/fixtures/statistics-input.json')).digest('hex'),
  ).toBe(oracle.inputSha256);
  expect(oracle.cases.map(({ ref, formula }) => ({ ref, formula }))).toEqual(spec.cases);
});
it.each(oracle.cases)('native Excel $ref $formula', ({ ref, value, kind }) => {
  const result = calculator([sheet]).result(sheet, ref);
  if (kind === 'error') expect(result).toEqual({ kind: 'error', value });
  else {
    expect(result.kind).toBe('value');
    if (result.kind === 'value') expect(result.value).toBeCloseTo(Number(value), 9);
  }
});
