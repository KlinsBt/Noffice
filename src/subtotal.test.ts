import { describe, expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import manual from '../tests/fixtures/subtotal-manual-input.json';
import filtered from '../tests/fixtures/subtotal-filtered-input.json';
import manualOracle from '../tests/fixtures/subtotal-manual-oracle.json';
import filteredOracle from '../tests/fixtures/subtotal-filtered-oracle.json';
import table from '../tests/fixtures/subtotal-table-input.json';
import tableOracle from '../tests/fixtures/subtotal-table-oracle.json';
import { calculator, formulaReferences } from './formulas';
import type { Sheet } from './model';

function fixture(spec = manual): Sheet {
  return {
    id: 'subtotal',
    name: spec.sheet,
    rows: 100,
    cols: 26,
    hiddenRows: spec !== manual ? filteredOracle.hiddenRows : manualOracle.hiddenRows,
    hiddenColumns: spec.hiddenColumns,
    definedNames: spec.names,
    ...(spec !== manual
      ? {
          filterMode: true,
          autoFilters: [{ ref: 'A1:B10', columns: [{ col: 1, values: ['keep'] }] }],
        }
      : {}),
    cells: Object.fromEntries([
      ...Object.entries(spec.cells).map(([ref, value]) => [
        ref,
        {
          value: String(value),
          ...(typeof value === 'string' && !value.startsWith('=') ? { dataType: 'text' } : {}),
        },
      ]),
      ...Object.entries(spec !== manual ? filtered.edits : {}).map(([ref, value]) => [
        ref,
        { value: String(value), dataType: 'text' },
      ]),
      ...spec.cases.map(({ ref, formula }) => [ref, { value: formula }]),
    ]),
  };
}
for (const [kind, spec, oracle] of [
  ['manual', manual, manualOracle],
  ['filtered', filtered, filteredOracle],
  ['table', table, tableOracle],
] as const) {
  describe(`native Excel SUBTOTAL ${kind}`, () => {
    it('matches the exact authored fixture and formulas', () => {
      expect(
        createHash('sha256')
          .update(readFileSync(`tests/fixtures/subtotal-${kind}-input.json`))
          .digest('hex'),
      ).toBe(oracle.inputSha256);
      expect(oracle.cases.map(({ ref, formula }) => ({ ref, formula }))).toEqual(spec.cases);
    });
    it.each(oracle.cases)('$ref $formula', ({ ref, value, kind }) => {
      const sheet = fixture(spec),
        actual = calculator([sheet]).result(sheet, ref);
      expect(actual.kind).toBe(kind);
      if (typeof value === 'number') expect(actual.value).toBeCloseTo(value, 12);
      else expect(actual.value).toBe(value);
    });
  });
}
it('uses persisted visibility for unsupported filter types without hiding formula errors', () => {
  const sheet = fixture(filtered);
  sheet.autoFilters![0].columns[0] = { col: 1, unsupported: true };
  sheet.cells.K1 = { value: '=IFERROR(SUBTOTAL(2,A2:A5),0)' };
  expect(calculator([sheet])(sheet, 'K1')).toBe(2);
  delete sheet.autoFilters;
  sheet.cells.K1 = { value: '=SUBTOTAL(2,K2)' };
  sheet.cells.K2 = { value: '=NOTIMPLEMENTED()' };
  expect(calculator([sheet]).result(sheet, 'K1').kind).toBe('unsupported');
});
it('keeps filter visibility until reapplication and treats visibility as a dynamic dependency', () => {
  const sheet = fixture(filtered);
  sheet.cells.K1 = { value: '=SUBTOTAL(9,A2:A5)' };
  expect(calculator([sheet])(sheet, 'K1')).toBe(10);
  sheet.cells.B4 = { value: 'keep', dataType: 'text' };
  expect(calculator([sheet])(sheet, 'K1')).toBe(10);
  sheet.cells.A5 = { value: '18' };
  expect(calculator([sheet])(sheet, 'K1')).toBe(20);
  expect(formulaReferences('=SUBTOTAL(109,A2:A5)+1')).toBeNull();
});
it('rejects scalar arguments and bounds sparse range traversal', () => {
  const sheet = fixture();
  sheet.cells.K1 = { value: '=SUBTOTAL(9,1)' };
  expect(calculator([sheet])(sheet, 'K1')).toBe('#VALUE!');
  sheet.cells = { Z10000: { value: '1' } };
  sheet.cells.K1 = { value: '=SUBTOTAL(9,A1:Z10000)' };
  expect(calculator([sheet]).result(sheet, 'K1').kind).toBe('unsupported');
});
