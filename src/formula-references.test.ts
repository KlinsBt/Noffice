import { describe, expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import spec from '../tests/fixtures/excel-formulas-input.json';
import oracle from '../tests/fixtures/excel-formulas-oracle.json';
import { calculator, formulaReferences, translateFormula } from './formulas';
import type { Sheet } from './model';

function fixture(): Sheet {
  const sheet: Sheet = {
    id: 'data',
    name: spec.sheet,
    rows: 100,
    cols: 26,
    cells: {},
    definedNames: spec.names,
    tables: [spec.table],
  };
  for (const [ref, value] of Object.entries(spec.cells))
    sheet.cells[ref] = {
      value: String(value),
      ...(typeof value === 'string' && !value.startsWith('=') ? { dataType: 'text' as const } : {}),
    };
  for (const { ref, formula } of spec.cases) sheet.cells[ref] = { value: formula };
  return sheet;
}

describe('independent desktop Excel formula results', () => {
  it('uses a receipt recorded for the exact authored input', () => {
    expect(
      createHash('sha256')
        .update(readFileSync('tests/fixtures/excel-formulas-input.json'))
        .digest('hex'),
    ).toBe(oracle.inputSha256);
    expect(oracle.cases.map(({ ref, formula }) => ({ ref, formula }))).toEqual(spec.cases);
  });
  it.each(oracle.cases)('$ref $formula', ({ ref, value, kind }) => {
    const sheet = fixture();
    expect(calculator([sheet]).result(sheet, ref)).toEqual({ value, kind });
  });
});

describe('reference safety and scope', () => {
  it('resolves sheet-local names and detects named cycles', () => {
    const a = fixture(),
      b: Sheet = {
        ...fixture(),
        id: 'other',
        name: 'Other',
        definedNames: { TaxRate: '0.4', loop: 'loop' },
        cells: {
          A1: { value: '=TaxRate' },
          A2: { value: "='Data Set'!TaxRate" },
          A3: { value: '=loop' },
        },
      };
    const calc = calculator([a, b]);
    expect(calc(b, 'A1')).toBe(0.4);
    expect(calc(b, 'A2')).toBe(0.2);
    expect(calc.result(b, 'A3').kind).toBe('unsupported');
  });
  it('uses the table containing an unqualified current-row formula', () => {
    const sheet = fixture();
    sheet.cells.D2 = { value: '=[@Amount]+[@Units]' };
    expect(calculator([sheet])(sheet, 'D2')).toBe(11);
  });
  it('distinguishes table names ending in digits from grid references', () => {
    const sheet = fixture();
    sheet.tables![0].name = 'Table1';
    sheet.definedNames!.TotalAmount = 'SUM(Table1[Amount])';
    sheet.cells.J2 = { value: '=SUM(Table1)' };
    sheet.cells.J3 = { value: '=TotalAmount' };
    const calc = calculator([sheet]);
    expect(calc(sheet, 'J2')).toBe(110);
    expect(calc(sheet, 'J3')).toBe(100);
    expect(translateFormula('=SUM(Table1)+A2', 1, 1)).toBe('=SUM(Table1)+B3');
  });
  it.each([
    '=IFNA(UNKNOWN(),0)',
    '=ISERROR(UNKNOWN())',
    '=IFNA(UNKNOWN(A:A),0)',
    '=ISNA(INDIRECT("J2"))',
  ])('never hides unsupported work or cycles: %s', (formula) => {
    const sheet = fixture();
    sheet.cells.J2 = { value: formula };
    expect(calculator([sheet]).result(sheet, 'J2').kind).toBe('unsupported');
  });
  it('flags indirect, named, table and whole-column dependencies for conservative recalculation', () => {
    for (const formula of [
      '=Gross',
      '=SUM(SalesData[Amount])',
      '=INDIRECT("B2")',
      '=VLOOKUP(1,A:B,2,0)',
    ])
      expect(formulaReferences(formula)).toBeNull();
    expect(formulaReferences('=SUM(A2:B3)')).toHaveLength(4);
  });
  it('does not translate cell-looking table column names during fill', () => {
    expect(translateFormula('=SalesData[A1]+B2', 1, 1)).toBe('=SalesData[A1]+C3');
  });
  it('does not invent values for external references or names with an unknown relative anchor', () => {
    const sheet = fixture();
    sheet.definedNames!.relative = 'B2';
    sheet.cells.J2 = { value: '=IFERROR(relative,0)' };
    sheet.cells.J3 = { value: "=IFERROR('[external.xlsx]Sheet1'!A1,0)" };
    const calc = calculator([sheet]);
    expect(calc.result(sheet, 'J2').kind).toBe('unsupported');
    expect(calc.result(sheet, 'J3').kind).toBe('unsupported');
  });
});
