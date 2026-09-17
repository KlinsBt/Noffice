import { describe, it, expect } from 'vitest';
import { calculator } from './formulas';
import type { Sheet } from './model';
const evaluate = (formula: string) => {
  const sheet: Sheet = {
    id: 's',
    name: 'Data',
    rows: 100,
    cols: 26,
    cells: {
      A1: { value: '10' },
      B1: { value: 'Ada', dataType: 'text' },
      C1: { value: '2' },
      A2: { value: '20' },
      B2: { value: 'Ben', dataType: 'text' },
      C2: { value: '3' },
      A3: { value: '30' },
      B3: { value: 'Casey', dataType: 'text' },
      C3: { value: '4' },
      D1: { value: formula },
    },
  };
  return calculator([sheet])(sheet, 'D1');
};
describe('lookup, criteria and typed calculation', () => {
  it.each([
    ['=VLOOKUP(20,A1:C3,2,FALSE)', 'Ben'],
    ['=VLOOKUP(25,A1:C3,2,TRUE)', 'Ben'],
    ['=VLOOKUP(5,A1:C3,2,TRUE)', '#N/A'],
    ['=VLOOKUP(25,A1:C3,2,FALSE)', '#N/A'],
    ['=VLOOKUP("20",A1:C3,2,FALSE)', '#N/A'],
    ['=VLOOKUP(20,A1:C3,4,FALSE)', '#REF!'],
    ['=VLOOKUP(20,A1:C3,0,FALSE)', '#VALUE!'],
    ['=VLOOKUP("B*",B1:C3,2,FALSE)', 3],
    ['=HLOOKUP("Ben",B2:C3,2,FALSE)', 'Casey'],
    ['=INDEX(A1:C3,2,2)', 'Ben'],
    ['=INDEX(A1:C3,4,2)', '#REF!'],
    ['=MATCH(20,A1:A3,0)', 2],
    ['=MATCH(25,A1:A3,1)', 2],
    ['=MATCH("ca*",B1:B3,0)', 3],
    ['=INDEX(B1:B3,MATCH(20,A1:A3,0))', 'Ben'],
    ['=COUNTIF(A1:A3,">=20")', 2],
    ['=COUNTIF(B1:B3,"B*")', 1],
    ['=SUMIF(A1:A3,">=20",C1:C3)', 7],
    ['=AVERAGEIF(A1:A3,">=20",C1:C3)', 3.5],
    ['=IFERROR(VLOOKUP(99,A1:C3,2,FALSE),"missing")', 'missing'],
    ['=IFERROR(FUTUREFUNCTION(A1),"wrong")', '#UNSUPPORTED!'],
    ['=ROUNDUP(1.234,2)', 1.24],
    ['=ROUNDDOWN(-1.239,2)', -1.23],
    ['=YEAR(DATE(2025,12,3))', 2025],
    ['=MONTH(DATE(2025,12,3))', 12],
    ['=DAY(DATE(2025,12,3))', 3],
    ['=TEXT(DATE(2025,12,3),"dd.mm.yyyy")', '03.12.2025'],
  ])('%s', (formula, result) => expect(evaluate(String(formula))).toBe(result));
});
