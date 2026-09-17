import { beforeAll, afterAll, it, expect } from 'vitest';
import JSZip from 'jszip';
import { normalizedTableHeaders, prepareHeaderEntry, applyHeaderEntry } from './sheet-header-entry';
import { importFile, exportOffice, nativeBackup } from './formats';
import { renameWorkbookTable } from './workbook-table-rename';
import { tableRenameFixture } from '../tests/fixtures/xlsx-table-rename';
import { calculator } from './formulas';
import { newFile, type WorkbookContent } from './model';
import { createSheetTable } from './sheet-tables';
const native = Blob.prototype.arrayBuffer;
beforeAll(() => {
  if (!native)
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
  if (!native) delete (Blob.prototype as Partial<Blob>).arrayBuffer;
});
it.each([
  [
    ['Region', 'Price', 'Price'],
    ['Region', 'Price', 'Price2'],
  ],
  [
    ['Region', 'Quantity', 'Quantity'],
    ['Region', 'Quantity', 'Quantity2'],
  ],
  [
    ['Quantity', 'Quantity', 'Quantity'],
    ['Quantity', 'Quantity2', 'Quantity3'],
  ],
  [
    ['Price', 'Price', 'Price2'],
    ['Price', 'Price3', 'Price2'],
  ],
  [
    ['Price', 'price', 'PRICE2'],
    ['Price', 'price3', 'PRICE2'],
  ],
  [
    ['Region', '', 'Price'],
    ['Region', 'Column1', 'Price'],
  ],
  [
    ['', '', ''],
    ['Column1', 'Column2', 'Column3'],
  ],
  [
    ['Column1', '', 'Column2'],
    ['Column1', 'Column12', 'Column2'],
  ],
  [
    [' label ', ' ', '123'],
    [' label ', ' ', '123'],
  ],
  [
    ['Quantity', 'Region', 'Price'],
    ['Quantity', 'Region', 'Price'],
  ],
])('normalizes the entire row, preserving explicit suffixes: %j', (values, expected) => {
  expect(normalizedTableHeaders(values)).toEqual(expected);
});
it('truncates headings and keeps duplicate suffixes within the 255-character limit', () => {
  expect(normalizedTableHeaders(['x'.repeat(260), 'x'.repeat(260)])).toEqual([
    'x'.repeat(255),
    'x'.repeat(254) + '2',
  ]);
});
async function fixture() {
  const input = await tableRenameFixture();
  const file = await importFile({
    name: 'Headers.xlsx',
    size: input.byteLength,
    arrayBuffer: async () => input,
  } as File);
  if (file.content.kind !== 'excel') throw Error('Expected workbook');
  const before = file.content;
  return {
    input,
    file,
    before,
    rename: (content: WorkbookContent, edit: Parameters<typeof renameWorkbookTable>[1]) =>
      renameWorkbookTable({ ...file, content }, edit),
  };
}
it('atomically repairs both colliding columns, caches and data edits while preserving unrelated package parts', async () => {
  const { file, before, input, rename } = await fixture();
  const proposed = structuredClone(before);
  proposed.sheets[0].cells.B1.value = 'Price';
  proposed.sheets[0].cells.B2.value = '4';
  const next = await applyHeaderEntry(before, proposed, rename),
    sheet = next.sheets[0];
  expect(before.sheets[0].cells.B1.value).toBe('Quantity');
  expect(proposed.sheets[0].cells.B1.value).toBe('Price');
  expect(sheet.tables![0].columns).toEqual(['Region', 'Price', 'Price2']);
  expect(sheet.cells.E1.value).toBe('=SUM(Sales[Price])');
  expect(sheet.cells.H4.value).toBe('=SUM(Sales[[Price]:[Price2]])');
  const calc = calculator(next.sheets);
  expect(calc(sheet, 'E1')).toBe(12);
  expect(calc(sheet, 'H4')).toBe(132);
  const zip = await JSZip.loadAsync(
      await (await exportOffice({ ...file, content: next })).arrayBuffer(),
    ),
    original = await JSZip.loadAsync(input);
  for (const path of [
    'xl/styles.xml',
    'xl/comments1.xml',
    'xl/theme/theme1.xml',
    'xl/worksheets/sheet2.xml',
    'custom/preservation.xml',
  ])
    expect(await zip.file(path)!.async('string')).toBe(await original.file(path)!.async('string'));
  const backup = nativeBackup({ ...file, content: next });
  const restored = await importFile({
    name: 'headers.noffice',
    size: backup.length,
    text: async () => backup,
  } as File);
  expect(restored.content).toEqual(next);
});
it('repairs swapped names simultaneously, including a calculated column and a header result', async () => {
  const { before, rename } = await fixture();
  const proposed = structuredClone(before);
  proposed.sheets[0].cells.A1.value = 'Quantity';
  proposed.sheets[0].cells.B1.value = 'Region';
  const next = await applyHeaderEntry(before, proposed, rename),
    s = next.sheets[0];
  expect(s.tables![0].columns).toEqual(['Quantity', 'Region', 'Price']);
  expect(s.cells.E1.value).toBe('=SUM(Sales[Region])');
  expect(calculator(next.sheets)(s, 'C2')).toBe(20);
  expect(s.cells.H3.cachedValue).toBe('Region');
});
it('handles unchanged pasted headings without changing their text type or breaking export', async () => {
  const { file, before, rename } = await fixture();
  const proposed = structuredClone(before);
  proposed.sheets[0].cells.B1.dataType = undefined;
  proposed.sheets[0].cells.B2.value = '4';
  const next = await applyHeaderEntry(before, proposed, rename);
  expect(next.sheets[0].cells.B1.dataType).toBe('text');
  expect((await exportOffice({ ...file, content: next })).size).toBeGreaterThan(0);
});
it('accepts apostrophe-prefixed literal formula headings but rejects live formulas and illegal controls atomically', async () => {
  const { before, rename } = await fixture();
  const proposed = structuredClone(before);
  proposed.sheets[0].cells.B1.value = "'=Heading";
  const next = await applyHeaderEntry(before, proposed, rename);
  expect(next.sheets[0].cells.B1).toMatchObject({ value: '=Heading', dataType: 'text' });
  for (const value of ['=1+2', 'bad\u0001name']) {
    proposed.sheets[0].cells.B1.value = value;
    expect(() => prepareHeaderEntry(before, proposed)).toThrow();
  }
  expect(before.sheets[0].cells.B1.value).toBe('Quantity');
});
it('rolls back a pasted header/data rectangle when any dependency is unsupported', async () => {
  const { file, before, input } = await fixture();
  const zip = await JSZip.loadAsync(input);
  zip.file('xl/connections.xml', '<connections/>');
  const data = await zip.generateAsync({ type: 'arraybuffer' });
  file.original!.data = data;
  const proposed = structuredClone(before);
  proposed.sheets[0].cells.B1.value = 'Units';
  proposed.sheets[0].cells.B2.value = '99';
  await expect(
    applyHeaderEntry(before, proposed, (content, edit) =>
      renameWorkbookTable({ ...file, content }, edit),
    ),
  ).rejects.toThrow(/not supported/);
  expect(before.sheets[0].cells.B1.value).toBe('Quantity');
  expect(before.sheets[0].cells.B2.value).toBe('2');
});
it('normalizes headers in a new workbook and keeps sheet creation available', async () => {
  const file = newFile('excel');
  if (file.content.kind !== 'excel') throw Error('Workbook');
  const s = file.content.sheets[0];
  s.cells = {
    A1: { value: 'Name' },
    B1: { value: 'Count' },
    A2: { value: 'Item' },
    B2: { value: '5' },
    D1: { value: '=SUM(Items[Count])' },
  };
  const before = createSheetTable(file.content, s.id, 'A1:B2', 'Items'),
    proposed = structuredClone(before);
  proposed.sheets[0].cells.A1.value = '';
  proposed.sheets[0].cells.B1.value = '';
  const next = await applyHeaderEntry(before, proposed, (content, edit) =>
    renameWorkbookTable({ ...file, content }, edit),
  );
  expect(next.sheets[0].tables![0].columns).toEqual(['Column1', 'Column2']);
  expect(next.xlsxStructureBase).toBeUndefined();
  expect(next.sheets[0].sourcePath).toBeUndefined();
  expect(calculator(next.sheets)(next.sheets[0], 'D1')).toBe(5);
});
