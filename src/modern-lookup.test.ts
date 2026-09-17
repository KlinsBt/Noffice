import { expect, it } from 'vitest';
import { calculator } from './formulas';
import type { Sheet } from './model';

const evaluate = (formula: string) => {
  const sheet: Sheet = { id: 's', name: 'Sheet1', rows: 100, cols: 26, cells: {} };
  for (const [row, key, result] of [
    [1, 10, 'Ada'],
    [2, 20, 'Ben'],
    [3, 20, 'Bea'],
    [4, 40, 'Casey'],
  ] as const) {
    sheet.cells['A' + row] = { value: String(key) };
    sheet.cells['B' + row] = { value: result, dataType: 'text' };
    sheet.cells['C' + row] = { value: String(50 - key) };
  }
  sheet.cells.D1 = { value: formula };
  return calculator([sheet]).result(sheet, 'D1');
};

it.each([
  ['=XLOOKUP(20,A1:A4,B1:B4)', 'Ben'],
  ['=_xlfn.XLOOKUP(20,A1:A4,B1:B4,,, -1)', 'Bea'],
  ['=XLOOKUP(25,A1:A4,B1:B4,,-1)', 'Ben'],
  ['=XLOOKUP(25,A1:A4,B1:B4,,1)', 'Casey'],
  ['=XLOOKUP(25,A1:A4,B1:B4,,1,2)', 'Casey'],
  ['=XLOOKUP(25,C1:C4,B1:B4,,1,-2)', 'Bea'],
  ['=XLOOKUP(25,C1:C4,B1:B4,,-1,-2)', 'Casey'],
  ['=XLOOKUP("B?a",B1:B4,A1:A4,,2)', 20],
  ['=XLOOKUP(10,A1:A4,B1:B4,1/0)', 'Ada'],
  ['=XLOOKUP(99,A1:A4,B1:B4,"missing")', 'missing'],
  ['=XMATCH(25,A1:A4,-1)', 2],
  ['=_xlfn.XMATCH(25,A1:A4,1)', 4],
  ['=XMATCH(20,A1:A4,0,-1)', 3],
  ['=XMATCH(5,A1:A4,1,2)', 1],
] as const)('scalar modern lookup: %s', (formula, value) =>
  expect(evaluate(formula)).toEqual({ value, kind: 'value' }),
);

it.each([
  ['=XLOOKUP(99,A1:A4,B1:B4)', '#N/A', 'error'],
  ['=XMATCH(99,A1:A4,0,2)', '#N/A', 'error'],
  ['=XLOOKUP(10,A1:A4,B1:B3)', '#VALUE!', 'error'],
  ['=XLOOKUP(10,A1:A4,B1:C4)', '#UNSUPPORTED!', 'unsupported'],
  ['=XMATCH(10,A1:A4,0,3)', '#VALUE!', 'error'],
] as const)('modern lookup failure: %s', (formula, value, kind) =>
  expect(evaluate(formula)).toEqual({ value, kind }),
);
