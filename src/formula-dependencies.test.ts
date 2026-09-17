import { expect, it } from 'vitest';
import { calculator } from './formulas';
import { FormulaDependencies } from './formula-dependencies';
import type { Sheet } from './model';
const sheet = (values: Record<string, string>, id = 'Data'): Sheet => ({
  id,
  name: id,
  rows: 100,
  cols: 26,
  cells: Object.fromEntries(Object.entries(values).map(([ref, value]) => [ref, { value }])),
});
const edit = (s: Sheet, ref: string, value: string): Sheet => ({
  ...s,
  cells: { ...s.cells, [ref]: { value } },
});

it('invalidates transitive cross-sheet dependents while retaining unrelated computed values', () => {
  let a = sheet({ A1: '2', B1: '=A1*3', Z9: '0' }),
    b = sheet({ A1: '=Data!B1+1', B1: '=2+2' }, 'Summary');
  let calc = calculator([a, b]);
  expect(calc(b, 'A1')).toBe(7);
  expect(calc(b, 'B1')).toBe(4);
  const before = calc.diagnostics().evaluations;
  const old = a;
  a = edit(a, 'A1', '5');
  calc = calc.update([a, b]);
  expect(calc(b, 'B1')).toBe(4);
  expect(calc.diagnostics().evaluations).toBe(before);
  expect(calc(b, 'A1')).toBe(16);
  expect(calc.diagnostics().evaluations).toBe(before + 2);
  expect(old.cells.A1.value).toBe('2');
  calc = calc.update([old, b]);
  expect(calc(b, 'A1')).toBe(7);
});
it('tracks missing aggregate cells, error dependencies and changed conditional branches', () => {
  let s = sheet({ A1: '1', B1: '=SUM(A1:A3)', C1: '=1/A2', D1: '=IF(A1=1,A2,A3)', Z9: '0' });
  let calc = calculator([s]);
  expect(calc(s, 'B1')).toBe(1);
  expect(calc.result(s, 'C1').kind).toBe('error');
  expect(calc(s, 'D1')).toBe(0);
  s = edit(s, 'A2', '2');
  calc = calc.update([s]);
  expect(calc(s, 'B1')).toBe(3);
  expect(calc(s, 'C1')).toBe(0.5);
  expect(calc(s, 'D1')).toBe(2);
  s = edit(edit(s, 'A1', '0'), 'A3', '8');
  calc = calc.update([s]);
  expect(calc(s, 'D1')).toBe(8);
  s = edit(s, 'A2', '4');
  calc = calc.update([s]);
  const before = calc.diagnostics().evaluations;
  expect(calc(s, 'D1')).toBe(8);
  expect(calc.diagnostics().evaluations).toBe(before);
});
it('invalidates names, extents, formula deletion and data-type changes', () => {
  let s = { ...sheet({ A1: '2', B1: '=Rate+A1', C1: '=SUM(A:A)' }), definedNames: { Rate: '3' } };
  let calc = calculator([s]);
  expect(calc(s, 'B1')).toBe(5);
  expect(calc(s, 'C1')).toBe(2);
  s = { ...s, definedNames: { Rate: '4' } };
  calc = calc.update([s]);
  expect(calc(s, 'B1')).toBe(6);
  s = { ...edit(s, 'A2', '10'), definedNames: s.definedNames };
  calc = calc.update([s]);
  expect(calc(s, 'C1')).toBe(12);
  s = { ...edit(s, 'B1', '9'), definedNames: s.definedNames };
  calc = calc.update([s]);
  expect(calc(s, 'B1')).toBe(9);
  s = { ...s, cells: { ...s.cells, B1: { value: '9', dataType: 'text' } } };
  calc = calc.update([s]);
  expect(calc(s, 'B1')).toBe('9');
});
it('invalidates all consumers of a shared lookup shortcut when a search key changes', () => {
  let s = sheet({
    A1: '1',
    A2: '2',
    B1: '10',
    B2: '20',
    D1: '=VLOOKUP(2,A1:B2,2,FALSE)',
    D2: '=VLOOKUP(2,A1:B2,2,FALSE)',
    Z9: '0',
  });
  let calc = calculator([s]);
  expect(calc(s, 'D1')).toBe(20);
  expect(calc(s, 'D2')).toBe(20);
  s = edit(s, 'A1', '2');
  calc = calc.update([s]);
  expect(calc(s, 'D1')).toBe(10);
  expect(calc(s, 'D2')).toBe(10);
});
it('handles dependency cycles and refuses to treat an overflowed graph as complete', () => {
  const graph = new FormulaDependencies(2);
  graph.read('B', 'A');
  graph.read('C', 'B');
  expect([...graph.affected(['A'])]).toEqual(['A', 'B', 'C']);
  graph.read('A', 'C');
  expect(graph.overflow).toBe(true);
  graph.clear();
  expect(graph.size).toBe(0);
  graph.read('A', 'B');
  graph.read('B', 'A');
  expect(graph.affected(['A']).size).toBe(2);
  let s = sheet({ A1: '=B1', B1: '=A1' });
  let calc = calculator([s]);
  expect(calc(s, 'A1')).toBe('#CYCLE!');
  s = edit(s, 'B1', '4');
  calc = calc.update([s]);
  expect(calc(s, 'A1')).toBe(4);
});
