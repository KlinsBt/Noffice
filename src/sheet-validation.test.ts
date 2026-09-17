import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { validationFixture } from '../tests/fixtures/xlsx-validation';
import {
  applyValidation,
  cellValidation,
  checkedValidation,
  type ValidationRule,
} from './sheet-validation';
import { validationError, validationChoices, validationSetupError } from './cell-validation';
import { importWorkbook, parseXML, elements } from './xlsx-import';
import { importFile, exportOffice } from './formats';
import type { WorkbookContent } from './model';
async function bytes(blob: Blob) {
  return new Promise<ArrayBuffer>((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as ArrayBuffer);
    r.onerror = reject;
    r.readAsArrayBuffer(blob);
  });
}
async function fixture() {
  const input = await validationFixture();
  const file = await importFile(
    Object.assign(new File([input], 'Validation.xlsx'), { arrayBuffer: async () => input }),
  );
  return { input, file, content: file.content as WorkbookContent };
}
const decimal: ValidationRule = {
  type: 'decimal',
  operator: 'between',
  formulae: ['0', '1'],
  allowBlank: true,
  showErrorMessage: true,
  errorStyle: 'stop',
  error: 'Between zero and one',
  prompt: 'Enter a fraction',
  showInputMessage: true,
};
it('rejects new rules whose bounds or list source cannot be resolved', async () => {
  const { content } = await fixture(),
    sheet = content.sheets[0];
  expect(
    validationSetupError({ ...decimal, formulae: ['BADFUNCTION()', '1'] }, sheet, content.sheets),
  ).toBeTruthy();
  expect(
    validationSetupError({ ...decimal, formulae: ['10', '1'] }, sheet, content.sheets),
  ).toBeTruthy();
  expect(
    validationSetupError(
      { ...decimal, type: 'list', formulae: ['MissingName'] },
      sheet,
      content.sheets,
    ),
  ).toBeTruthy();
  expect(validationSetupError(decimal, sheet, content.sheets)).toBeNull();
});
it('resolves validation on blank cells and full-column rules without expanding a million cells', async () => {
  const { content } = await fixture();
  const sheet = content.sheets[0];
  expect(Object.keys(sheet.cells).length).toBeLessThan(20);
  expect(cellValidation(sheet, 'A3')?.formulae).toEqual(['B3', 'B3+10']);
  expect(cellValidation(sheet, 'D9000')?.type).toBe('list');
  expect(
    validationChoices(
      { value: '', validation: cellValidation(sheet, 'D9000') },
      sheet,
      content.sheets,
    ),
  ).toEqual(['Red', 'Green']);
  expect(validationError({ value: '' }, '2', sheet, content.sheets, 'A3')).toBeTruthy();
  expect(validationError({ value: '' }, '3', sheet, content.sheets, 'A3')).toBeNull();
  expect(validationError({ value: '' }, '=B3+2', sheet, content.sheets, 'A3')).toBeNull();
});
it('applies range rules atomically, translates relative formulas and leaves existing invalid values untouched', async () => {
  const { content } = await fixture();
  const sheet = content.sheets[0];
  const next = applyValidation(sheet, ['A2', 'A3'], { ...decimal, formulae: ['B2', 'B2+1'] });
  expect(next.cells.A2.value).toBe('5');
  expect(sheet.cells.A2.validation?.type).toBe('whole');
  expect(next.cells.A3.validation?.formulae).toEqual(['B3', 'B3+1']);
  expect(() => applyValidation(content.sheets[2], ['A1'], decimal)).toThrow(/protected/);
  expect(() => checkedValidation({ ...decimal, formulae: [''] })).toThrow();
  expect(() => checkedValidation({ ...decimal, errorStyle: 'warning' })).toThrow();
});
it.each([
  ['between', '0.5', true],
  ['between', '2', false],
  ['notBetween', '2', true],
  ['equal', '0', true],
  ['notEqual', '0', false],
  ['greaterThan', '0.5', true],
  ['lessThan', '-1', true],
  ['greaterThanOrEqual', '0', true],
  ['lessThanOrEqual', '0', true],
])('enforces %s for %s', async (operator, value, valid) => {
  const sheet = { id: 'a', name: 'a', rows: 100, cols: 26, cells: {} };
  expect(
    validationError(
      { value: '', validation: { ...decimal, operator: String(operator) } },
      String(value),
      sheet,
      [sheet],
    ) === null,
  ).toBe(valid);
});
it('patches a range member, shifts surviving formulas and keeps unrelated package payloads', async () => {
  const { input, file, content } = await fixture();
  content.sheets[0] = applyValidation(content.sheets[0], ['A2'], decimal);
  content.sheets[0] = applyValidation(content.sheets[0], ['D2'], {
    ...decimal,
    type: 'list',
    formulae: ['"Blue,Orange"'],
  });
  content.sheets[0] = applyValidation(content.sheets[0], ['D4'], { type: 'none', formulae: [] });
  const output = await bytes(await exportOffice(file)),
    read = await importWorkbook(output);
  expect(cellValidation(read.sheets[0], 'A2')?.type).toBe('decimal');
  expect(cellValidation(read.sheets[0], 'A3')?.formulae).toEqual(['B3', 'B3+10']);
  expect(cellValidation(read.sheets[0], 'D3')?.formulae).toEqual(['"Red,Green"']);
  expect(cellValidation(read.sheets[0], 'D4')).toBeUndefined();
  expect(cellValidation(read.sheets[0], 'D9000')?.formulae).toEqual(['"Red,Green"']);
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && !['xl/worksheets/sheet1.xml', 'xl/workbook.xml'].includes(path))
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const xml = parseXML(await after.file('xl/worksheets/sheet1.xml')!.async('string'));
  expect(
    elements(xml, 'dataValidation')
      .find((n) => n.getAttribute('sqref') === 'A3:A3')
      ?.getAttribute('imeMode'),
  ).toBe('noControl');
});
it('writes new-workbook validation with the same range and cell overrides', async () => {
  const { file, content } = await fixture();
  delete file.original;
  content.sheets[0] = applyValidation(content.sheets[0], ['D2'], decimal);
  const read = await importWorkbook(await bytes(await exportOffice(file)));
  expect(cellValidation(read.sheets[0], 'D2')?.type).toBe('decimal');
  expect(cellValidation(read.sheets[0], 'D3')?.type).toBe('list');
  expect(cellValidation(read.sheets[0], 'A3')?.formulae).toEqual(['B3', 'B3+10']);
});
