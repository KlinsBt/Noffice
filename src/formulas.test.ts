import { describe, expect, it } from 'vitest';
import {
  address,
  calculator,
  colName,
  coordinates,
  displayValue,
  rangeAddresses,
  translateFormula,
} from './formulas';
import type { Sheet } from './model';
const sheet = (cells: Record<string, string>, name = 'Sheet 1'): Sheet => ({
  id: name,
  name,
  rows: 100,
  cols: 26,
  cells: Object.fromEntries(Object.entries(cells).map(([key, value]) => [key, { value }])),
});
function calculate(formula: string, cells: Record<string, string> = {}) {
  const s = sheet({ ...cells, Z99: formula });
  return calculator([s])(s, 'Z99');
}
describe('formula evaluation', () => {
  it.each([
    ['=2+3*4', 14],
    ['=(2+3)*4', 20],
    ['=2^3^2', 64],
    ['=-2^2', 4],
    ['=25%*200', 50],
    ['=SUM(1,2,3)', 6],
    ['=AVERAGE(2,4,6)', 4],
    ['=MIN(3,1,2)', 1],
    ['=MAX(3,1,2)', 3],
    ['=ROUND(1.235,2)', 1.24],
    ['=ROUND(-1.25,1)', -1.3],
    ['=ABS(-5)', 5],
    ['=INT(-1.2)', -2],
    ['=SQRT(81)', 9],
    ['=MOD(-3,2)', 1],
    ['=POWER(3,3)', 27],
    ['=COUNT(1,"a",3)', 2],
    ['=COUNTA(1,"a","")', 3],
    ['=IF(1<2,"yes",1/0)', 'yes'],
    ['=IF(FALSE,1/0,7)', 7],
    ['=IFERROR(1/0,42)', 42],
    ['=AND(TRUE,1>0)', true],
    ['=OR(FALSE,TRUE)', true],
    ['=NOT(FALSE)', true],
    ['="Hi"&" there"', 'Hi there'],
    ['=CONCAT("A","B")', 'AB'],
    ['=LEN("hello")', 5],
    ['=UPPER("hi")', 'HI'],
    ['=LOWER("HELLO")', 'hello'],
    ['=TRIM("  one   two  ")', 'one two'],
    ['=LEFT("hello",2)', 'he'],
    ['=RIGHT("hello",2)', 'lo'],
    ['=RIGHT("hello",0)', ''],
    ['=1/0', '#DIV/0!'],
    ['=SUM(', '#ERROR!'],
    ['=UNKNOWN(1)', '#UNSUPPORTED!'],
    ['=AVERAGE("a")', '#VALUE!'],
    ['=SQRT(-1)', '#NUM!'],
    ['=1+"hi"', '#VALUE!'],
    ['=IF(1)', '#VALUE!'],
    ['=window.alert(1)', '#UNSUPPORTED!'],
  ])('%s => %s', (formula, result) => expect(calculate(String(formula))).toBe(result));
  it('recalculates nested references and ranges', () =>
    expect(calculate('=SUM(B1:B3)', { A1: '5', B1: '=A1*2', B2: '=B1+1', B3: '=B2+1' })).toBe(33));
  it('does not treat numeric strings in referenced text cells as numbers', () =>
    expect(calculate('=SUM(A1:A3)', { A1: '2', A2: "'12", A3: '3' })).toBe(5));
  it('reports cycles and allows lazy IF to skip one', () => {
    expect(calculate('=A1', { A1: '=B1', B1: '=A1' })).toBe('#CYCLE!');
    expect(calculate('=IF(FALSE,Z99,5)')).toBe(5);
  });
  it('evaluates quoted cross-sheet references and errors for missing sheets', () => {
    const a = sheet({ A1: "=SUM('Other sheet'!A1:A2)" }),
      b = sheet({ A1: '3', A2: '7' }, 'Other sheet');
    expect(calculator([a, b])(a, 'A1')).toBe(10);
    expect(calculator([a])(a, 'A1')).toBe('#REF!');
  });
  it('skips the empty tail of large aggregate references', () =>
    expect(calculate('=SUM(A1:A999999)')).toBe(0));
  it('bounds traversal across a sparse but distant used cell', () =>
    expect(calculate('=SUM(A1:A999999)', { A999999: '1' })).toBe('#LIMIT!'));
});
describe('grid primitives', () => {
  it('round trips addresses beyond Z', () => {
    expect(colName(26)).toBe('AA');
    expect(address(99, 255)).toBe('IV100');
    expect(coordinates('$IV$100')).toEqual([99, 255]);
  });
  it('normalizes reversed range selection', () =>
    expect(rangeAddresses('B2', 'A1')).toEqual(['A1', 'B1', 'A2', 'B2']));
  it('translates relative but preserves absolute refs and quoted strings', () =>
    expect(translateFormula('=A1+$B1+C$2+$D$4+"A1"', 2, 1)).toBe('=B3+$B3+D$2+$D$4+"A1"'));
  it('does not translate a function named LOG10', () =>
    expect(translateFormula('=LOG10(A1)', 1, 0)).toBe('=LOG10(A2)'));
  it('reports invalid translated refs', () =>
    expect(translateFormula('=A1', -1, 0)).toBe('=#REF!'));
  it('formats currencies and percentages without altering values', () => {
    expect(displayValue(1234.5, { value: '1234.5', format: 'currency' })).toBe('$1,234.50');
    expect(displayValue(0.25, { value: '.25', format: 'percent' })).toBe('25%');
  });
});
