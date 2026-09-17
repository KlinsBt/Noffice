import { describe, it, expect } from 'vitest';
import type { Cell, Sheet } from './model';
import { validationChoices, validationError } from './cell-validation';
const a: Sheet = {
  id: 'a',
  name: 'Input',
  rows: 100,
  cols: 26,
  cells: {},
  definedNames: { choices: "'List values'!$A$1:$A$2" },
};
const b: Sheet = {
  id: 'b',
  name: 'List values',
  rows: 100,
  cols: 26,
  cells: { A1: { value: 'First' }, A2: { value: 'Second' } },
};
describe('worksheet validation rules', () => {
  it('resolves a named list on another sheet without fetching any links', () => {
    const cell: Cell = {
      value: '',
      validation: { type: 'list', formulae: ['choices'], showErrorMessage: true },
    };
    expect(validationChoices(cell, a, [a, b])).toEqual(['First', 'Second']);
    expect(validationError(cell, 'Third', a, [a, b])).toBeTruthy();
    expect(validationError(cell, 'first', a, [a, b])).toBeNull();
  });
  it('enforces whole-number stop rules but respects optional input and warning-only rules', () => {
    const cell: Cell = {
      value: '',
      validation: {
        type: 'whole',
        operator: 'between',
        formulae: [1, 10],
        showErrorMessage: true,
        allowBlank: true,
      },
    };
    expect(validationError(cell, '1.5', a, [a])).toBeTruthy();
    expect(validationError(cell, '11', a, [a])).toBeTruthy();
    expect(validationError(cell, '4', a, [a])).toBeNull();
    expect(validationError(cell, '', a, [a])).toBeNull();
    cell.validation!.errorStyle = 'warning';
    expect(validationError(cell, '11', a, [a])).toBeNull();
  });
  it('leaves external list definitions unresolved rather than opening an external workbook', () => {
    const cell: Cell = {
      value: '',
      validation: { type: 'list', formulae: ["'[1]Remote'!$A$1:$A$9"] },
    };
    expect(validationChoices(cell, a, [a, b])).toBeNull();
  });
});
