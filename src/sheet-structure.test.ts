import { expect, it, beforeAll, afterAll } from 'vitest';
import JSZip from 'jszip';
import { structureFixture } from '../tests/fixtures/xlsx-structure';
import {
  checkedStructureEdit,
  structuralFormula,
  structuralRange,
  type StructureEdit,
} from './sheet-structure';
import { restructureXlsx } from './xlsx-structure';
import { importWorkbook, parseXML, elements } from './xlsx-import';
import { importFile, exportOffice, nativeBackup } from './formats';
import { editWorkbookStructure } from './workbook-structure';
import type { WorkbookContent } from './model';
import { newFile } from './model';
import { calculator } from './formulas';
import { cellValidation } from './sheet-validation';

const blobArrayBuffer = Blob.prototype.arrayBuffer;
beforeAll(() => {
  if (!blobArrayBuffer)
    Blob.prototype.arrayBuffer = function () {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as ArrayBuffer);
        reader.onerror = reject;
        reader.readAsArrayBuffer(this);
      });
    };
});
afterAll(() => {
  if (!blobArrayBuffer) delete (Blob.prototype as Partial<Blob>).arrayBuffer;
});

const insert: StructureEdit = { sheet: 'Data', axis: 'row', action: 'insert', at: 2, count: 2 };
const del: StructureEdit = { ...insert, action: 'delete' };
it.each([
  ['A1', 'A1', 'A1'],
  ['A3', 'A5', '#REF!'],
  ['$A$5', '$A$7', '$A$3'],
  ['A1:A5', 'A1:A7', 'A1:A3'],
  ['A3:A4', 'A5:A6', '#REF!'],
  ['A2:A3', 'A2:A5', 'A2:A2'],
  ['A4:A6', 'A6:A8', 'A3:A4'],
  ['A6:A2', 'A8:A2', 'A4:A2'],
  ['2:6', '2:8', '2:4'],
  ['A:C', 'A:C', 'A:C'],
  ['A1:A1048576', 'A1:A1048576', 'A1:A1048574'],
])('repairs range %s for insertion and deletion', (source, inserted, deleted) => {
  expect(structuralRange(source, insert)).toBe(inserted);
  expect(structuralRange(source, del)).toBe(deleted);
});
it('repairs column ranges and absolute mixed references without rewriting strings, names or functions', () => {
  const op = { ...insert, axis: 'column' as const, at: 1, count: 1 };
  expect(
    structuralFormula('=SUM($A2:C$5)+LOG10(100)+A3_name+"B3"+INDIRECT("B3")+SUM(B:D)', 'Data', op),
  ).toBe('=SUM($A2:D$5)+LOG10(100)+A3_name+"B3"+INDIRECT("B3")+SUM(C:E)');
  expect(structuralFormula('=Data!$A$3+Other!A3+A3+"Data!A3"', 'Other', insert)).toBe(
    '=Data!$A$5+Other!A3+A3+"Data!A3"',
  );
  expect(
    structuralFormula("='O''Brien'!$B$4+Sheet2!B4", 'Other', { ...insert, sheet: "O'Brien" }),
  ).toBe("='O''Brien'!$B$6+Sheet2!B4");
  expect(structuralFormula('=Data!A3+Data!A2:A5', 'Other', del)).toBe('=Data!#REF!+Data!A2:A3');
});
it.each([
  '=SUM(Sheet1:Sheet3!A1)',
  "='Sheet 1:Sheet 3'!A1",
  '=Table1[Amount]',
  "='[Book.xlsx]Data'!A1",
])('rejects unsupported reference syntax atomically: %s', (formula) =>
  expect(() => structuralFormula(formula, 'Data', insert)).toThrow(),
);
it('bounds operations and respects identifiers outside the A1 grid', () => {
  expect(() => checkedStructureEdit({ ...insert, count: -1 })).toThrow();
  expect(() => checkedStructureEdit({ ...insert, at: 1048576 })).toThrow();
  expect(structuralFormula('=XFE1+AAA1234567', 'Data', insert)).toBe('=XFE1+AAA1234567');
  expect(structuralFormula('=SUM(A2 : A5)', 'Data', del)).toBe('=SUM(A2:A3)');
  expect(structuralFormula('=LOG10 (100)+SUM(A2 : A5)', 'Data', del)).toBe(
    '=LOG10 (100)+SUM(A2:A3)',
  );
  expect(structuralFormula("=Data ! A3+'Data' ! $A$5+A5", 'Other', insert)).toBe(
    "=Data ! A5+'Data' ! $A$7+A5",
  );
  expect(() => structuralFormula('=SUM(Data!A2:Data!A5)', 'Data', del)).toThrow();
});
it('supports new workbooks and multiple selected rows without requiring an imported original', async () => {
  const file = newFile('excel');
  const book = file.content as WorkbookContent;
  book.sheets[0].cells = { A1: { value: '2' }, A2: { value: '3' }, B1: { value: '=SUM(A1:A2)' } };
  file.content = await editWorkbookStructure(file, {
    ...insert,
    sheet: 'Sheet 1',
    at: 1,
    count: 3,
  });
  expect(file.original).toBeUndefined();
  expect(file.content.xlsxStructureBase).toBeUndefined();
  expect(file.content.sheets[0].sourcePath).toBeUndefined();
  expect(file.content.sheets[0].cells.A5.value).toBe('3');
  expect(file.content.sheets[0].cells.B1.value).toBe('=SUM(A1:A5)');
  const next = await importWorkbook(await (await exportOffice(file)).arrayBuffer());
  expect(next.sheets[0].cells.B1.value).toBe('=SUM(A1:A5)');
  expect(calculator(next.sheets)(next.sheets[0], 'B1')).toBe(5);
  file.content.sheets.push({
    id: 'added',
    name: 'Added sheet',
    cells: { A1: { value: 'Still editable' } },
    rows: 100,
    cols: 26,
  });
  const added = await importWorkbook(await (await exportOffice(file)).arrayBuffer());
  expect(added.sheets[1].cells.A1.value).toBe('Still editable');
});
it('inherits validation at the preceding range edge and translates relative copied bounds', async () => {
  const zip = await JSZip.loadAsync(await structureFixture());
  const doc = parseXML(await zip.file('xl/worksheets/sheet1.xml')!.async('string'));
  elements(doc, 'formula1')[0].textContent = '$B$2';
  elements(doc, 'formula2')[0].textContent = 'B2+10';
  zip.file('xl/worksheets/sheet1.xml', new XMLSerializer().serializeToString(doc));
  const input = await zip.generateAsync({ type: 'arraybuffer' });
  const columns = await importWorkbook(
    await restructureXlsx(input, [{ ...insert, axis: 'column', at: 1, count: 1 }]),
  );
  expect(columns.sheets[0].validationRanges?.[0].ref).toBe('A2:B6');
  expect(cellValidation(columns.sheets[0], 'B2')?.formulae).toEqual(['$C$2', 'D2+10']);
  const rows = await importWorkbook(await restructureXlsx(input, [{ ...insert, at: 6, count: 1 }]));
  expect(rows.sheets[0].validationRanges?.[0].ref).toBe('A2:A7');
  expect(cellValidation(rows.sheets[0], 'A7')?.formulae).toEqual(['$B$2', 'B7+10']);
});
it('rejects pushing a represented cell beyond the editable grid', async () => {
  const zip = await JSZip.loadAsync(await structureFixture());
  const doc = parseXML(await zip.file('xl/worksheets/sheet1.xml')!.async('string'));
  const cell = doc.createElementNS(doc.documentElement.namespaceURI, 'c');
  cell.setAttribute('r', 'IV1');
  elements(doc, 'row')[0].appendChild(cell);
  zip.file('xl/worksheets/sheet1.xml', new XMLSerializer().serializeToString(doc));
  await expect(
    restructureXlsx(await zip.generateAsync({ type: 'arraybuffer' }), [
      { ...insert, axis: 'column', at: 0, count: 1 },
    ]),
  ).rejects.toThrow('256 columns');
});
it('moves styles, shared formulas, merges, dimensions, filters, validation, names and print references', async () => {
  const input = await structureFixture();
  const output = await restructureXlsx(input, [insert]);
  const book = await importWorkbook(output),
    sheet = book.sheets[0];
  expect(sheet.cells.A5.value).toBe('20');
  expect(sheet.cells.B2.value).toBe('=A2+$A$6');
  expect(sheet.cells.B8.value).toBe('=A8*2');
  expect(sheet.cells.A3.bold).toBe(true);
  expect(sheet.cells.A4.fill).toBe('#FFE699');
  expect(sheet.rowHeights?.[2]).toBe(36);
  expect(sheet.hiddenRows).toContain(6);
  expect(sheet.merges).toEqual(['E5:F5', 'E7:F8']);
  expect(sheet.validationRanges?.[0].ref).toBe('A2:A8');
  expect(sheet.autoFilters?.[0].ref).toBe('A1:B8');
  expect(sheet.definedNames?.amounts).toBe('Data!$A$2:$A$8');
  expect(book.sheets[1].cells.A1.value).toBe('=Data!A5');
  expect(book.sheets[1].cells.A5.hyperlink).toBe("#'Data'!$A$5");
  const zip = await JSZip.loadAsync(output),
    source = await JSZip.loadAsync(input);
  expect(await zip.file('xl/workbook.xml')!.async('string')).toContain("'Data'!$A1:$F8");
  for (const path of [
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/worksheets/sheet3.xml',
    'custom/preservation.xml',
  ])
    expect(await zip.file(path)!.async('uint8array')).toEqual(
      await source.file(path)!.async('uint8array'),
    );
});
it('deletes ranges and repairs cross-sheet #REF errors without moving the source input', async () => {
  const input = await structureFixture(),
    copy = new Uint8Array(input).slice();
  const book = await importWorkbook(await restructureXlsx(input, [del]));
  expect(book.sheets[0].cells.A3.value).toBe('40');
  expect(book.sheets[0].cells.B2.value).toBe('=A2+#REF!');
  expect(book.sheets[1].cells.A1.value).toBe('=Data!#REF!');
  expect(calculator(book.sheets)(book.sheets[1], 'A1')).toBe('#REF!');
  expect(book.sheets[0].merges).toEqual(['E3:F4']);
  expect(new Uint8Array(input)).toEqual(copy);
});
it('keeps edits across consecutive checkpoints, backup reload, subsequent exports and undo snapshots', async () => {
  const input = await structureFixture();
  const file = await importFile(
    Object.assign(new File([input], 'Structure.xlsx'), { arrayBuffer: async () => input }),
  );
  const original = file.content as WorkbookContent;
  original.sheets[0].cells.A2.value = '12';
  original.sheets[0].cells.A2.validation = {
    type: 'decimal',
    operator: 'between',
    formulae: ['0', '50'],
    showErrorMessage: true,
  };
  file.content = await editWorkbookStructure(file, insert);
  const checkpoint = file.content;
  file.content.sheets[0].cells.A5.value = '25';
  file.content = await editWorkbookStructure(file, { ...insert, axis: 'column', at: 1, count: 1 });
  expect(file.original?.data).toBe(input);
  expect(file.content.sheets[0].cells.A5.value).toBe('25');
  expect(file.content.sheets[0].cells.C2.value).toBe('=A2+$A$6');
  expect(file.content.sheets[0].cells.A2.validation?.type).toBe('decimal');
  const backup = nativeBackup(file);
  const restored = await importFile(
    Object.assign(new File([backup], 'Structure.noffice'), { text: async () => backup }),
  );
  const result = await importWorkbook(await (await exportOffice(restored)).arrayBuffer());
  expect(result.sheets[0].cells.C2.value).toBe('=A2+$A$6');
  expect(result.sheets[0].cells.A2.validation?.type).toBe('decimal');
  file.content = checkpoint;
  expect(
    (await importWorkbook(await (await exportOffice(file)).arrayBuffer())).sheets[0].cells.B2.value,
  ).toBe('=A2+$A$6');
});
it.each(['sheetProtection', 'drawing', 'legacyDrawing', 'extLst'])(
  'rejects unsupported %s before returning an edited package',
  async (tag) => {
    const zip = await JSZip.loadAsync(await structureFixture());
    const doc = parseXML(await zip.file('xl/worksheets/sheet1.xml')!.async('string'));
    doc.documentElement.appendChild(doc.createElementNS(doc.documentElement.namespaceURI, tag));
    zip.file('xl/worksheets/sheet1.xml', new XMLSerializer().serializeToString(doc));
    await expect(
      restructureXlsx(await zip.generateAsync({ type: 'arraybuffer' }), [insert]),
    ).rejects.toThrow();
  },
);
