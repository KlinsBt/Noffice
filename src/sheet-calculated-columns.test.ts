import { it, expect, beforeAll, afterAll } from 'vitest';
import JSZip from 'jszip';
import { newFile, type WorkbookContent } from './model';
import { calculatedColumnEntry } from './sheet-calculated-columns';
import { createSheetTable, editSheetTable, defaultTableStyle } from './sheet-tables';
import { calculator } from './formulas';
import { importFile, exportOffice, nativeBackup } from './formats';
import { importWorkbook } from './xlsx-import';
import { hydrateTableMetadata } from './workbook-tables';
import { renameWorkbookTable } from './workbook-table-rename';
import { calculatedColumnsFixture } from '../tests/fixtures/xlsx-calculated-columns';
import { tableRenameFixture } from '../tests/fixtures/xlsx-table-rename';
const native = Blob.prototype.arrayBuffer;
beforeAll(() => {
  if (!native)
    Blob.prototype.arrayBuffer = function () {
      return new Promise((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(r.result as ArrayBuffer);
        r.onerror = reject;
        r.readAsArrayBuffer(this);
      });
    };
});
afterAll(() => {
  if (!native) delete (Blob.prototype as Partial<Blob>).arrayBuffer;
});
function fresh() {
  const file = newFile('excel'),
    content = file.content as WorkbookContent,
    s = content.sheets[0];
  s.cells = {
    A1: { value: 'Quantity' },
    B1: { value: 'Total' },
    A2: { value: '2' },
    A3: { value: '3' },
    A4: { value: '5' },
    D1: { value: '=SUM(Sales[Total])' },
  };
  return createSheetTable(content, s.id, 'A1:B4', 'Sales');
}
function entry(content: WorkbookContent, ref: string, value: string) {
  return calculatedColumnEntry(content, content.sheets[0].id, ref, value)!;
}
function values(content: WorkbookContent) {
  const s = content.sheets[0],
    calc = calculator(content.sheets);
  return ['B2', 'B3', 'B4', 'D1'].map((ref) => calc(s, ref));
}
it('fills above and below an entry, preserves mixed locks and styles, and leaves the original untouched', () => {
  const before = fresh();
  before.sheets[0].hiddenRows = [3];
  before.sheets[0].cells.B4 = { value: '', bold: true, numFmt: '0.00', note: 'Keep' };
  const next = entry(before, 'B3', '=A3*10+$A$2');
  expect(values(next)).toEqual([22, 32, 52, 106]);
  expect(next.sheets[0].cells.B4).toMatchObject({
    value: '=A4*10+$A$2',
    bold: true,
    numFmt: '0.00',
    note: 'Keep',
  });
  expect(next.sheets[0].tables![0].calculatedColumns).toEqual([null, 'A2*10+$A$2']);
  expect(before.sheets[0].cells.B3).toBeUndefined();
  expect(before.sheets[0].cells.B4.value).toBe('');
});
it('creates and updates a uniform structured calculated column', () => {
  const first = entry(fresh(), 'B3', '=Sales[[#This Row],[Quantity]]*10');
  expect(values(first)).toEqual([20, 30, 50, 100]);
  const next = entry(first, 'B4', '=Sales[[#This Row],[Quantity]]*20');
  expect(values(next)).toEqual([40, 60, 100, 200]);
  expect(values(first)).toEqual([20, 30, 50, 100]);
});
it.each(['99', '', '=A4*2'])('does not overwrite a calculated-column exception %j', (exception) => {
  const content = entry(fresh(), 'B2', '=A2*10');
  content.sheets[0].cells.B4.value = exception;
  expect(entry(content, 'B3', '=A3*20')).toBeUndefined();
  expect(content.sheets[0].tables![0].calculatedColumns![1]).toBe('A2*10');
});
it('does not replace existing data elsewhere, but can replace the only existing value', () => {
  const content = fresh();
  content.sheets[0].cells.B2 = { value: '99' };
  expect(entry(content, 'B3', '=A3*10')).toBeUndefined();
  expect(values(entry(content, 'B2', '=A2*10'))).toEqual([20, 30, 50, 100]);
});
it('keeps text that resembles a formula as an exception', () => {
  const content = entry(fresh(), 'B2', '=A2*10');
  content.sheets[0].cells.B4.dataType = 'text';
  expect(entry(content, 'B3', '=A3*20')).toBeUndefined();
});
it.each([
  ['B1', '=1'],
  ['D3', '=1'],
  ['B3', '42'],
  ['B3', "'=A3"],
])('leaves ordinary entry %s %s to the cell editor', (ref, value) => {
  expect(entry(fresh(), ref, value)).toBeUndefined();
});
it('excludes a totals row and supports a single body row', () => {
  const content = fresh(),
    s = content.sheets[0];
  s.tables![0].totalRows = 1;
  s.cells.B4 = { value: '=SUM(B2:B3)' };
  expect(entry(content, 'B4', '=1')).toBeUndefined();
  expect(entry(content, 'B2', '=A2*10').sheets[0].cells.B4.value).toBe('=SUM(B2:B3)');
  s.tables![0].ref = 'A1:B2';
  s.tables![0].totalRows = 0;
  expect(entry(content, 'B2', '=A2*10').sheets[0].tables![0].calculatedColumns![1]).toBe('A2*10');
});
it('rejects protected, merged, overlarge and extended columns before changing any cell', () => {
  for (const kind of ['protected', 'merged', 'limit', 'extended', 'workbook']) {
    const content = fresh(),
      s = content.sheets[0];
    if (kind === 'protected') s.protected = true;
    if (kind === 'merged') s.merges = ['B3:C3'];
    if (kind === 'limit') s.tables![0].ref = 'A1:B10001';
    if (kind === 'extended') s.tables![0].calculationBlocked = 'Extended metadata';
    if (kind === 'workbook') s.tableEditingBlocked = 'Workbook protection';
    expect(() => entry(content, 'B3', '=A3*10')).toThrow();
    expect(s.cells.B3).toBeUndefined();
  }
});
it('rejects a failed validation anywhere in the fill atomically', () => {
  const content = fresh();
  content.sheets[0].cells.B4 = {
    value: '',
    validation: {
      type: 'whole',
      operator: 'lessThan',
      formulae: [40],
      showErrorMessage: true,
      error: 'Too large',
    },
  };
  expect(() => entry(content, 'B3', '=A3*10')).toThrow('B4: Too large');
  expect(() => entry(content, 'B3', '=Sales[[#This Row],[Quantity]]*10')).toThrow('B4: Too large');
  expect(content.sheets[0].cells.B3).toBeUndefined();
});
async function fixture(input = calculatedColumnsFixture()) {
  const bytes = await input;
  const file = await importFile({
    name: 'Calculated.xlsx',
    size: bytes.byteLength,
    arrayBuffer: async () => bytes,
  } as File);
  if (file.content.kind !== 'excel') throw Error('Workbook');
  return { file, content: file.content, bytes };
}
it('exports retained master formulas, cells and dependent caches without changing unrelated ZIP parts or originals', async () => {
  const { file, content, bytes } = await fixture();
  const next = entry(content, 'C3', '=B3*10+$B$2');
  const output = await (await exportOffice({ ...file, content: next })).arrayBuffer();
  const zip = await JSZip.loadAsync(output),
    old = await JSZip.loadAsync(bytes);
  expect(await zip.file('xl/tables/table1.xml')!.async('string')).toContain(
    '<calculatedColumnFormula>B2*10+$B$2</calculatedColumnFormula>',
  );
  expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toMatch(
    /<c r="E1"[^>]*><f>SUM\(Sales\[Price\]\)<\/f><v>106<\/v>/,
  );
  for (const path of [
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/comments1.xml',
    'xl/worksheets/sheet2.xml',
    'custom/preservation.xml',
  ])
    expect(await zip.file(path)!.async('string')).toBe(await old.file(path)!.async('string'));
  const parsed = await importWorkbook(output);
  expect(parsed.sheets[0].tables![0].calculatedColumns).toEqual([null, null, 'B2*10+$B$2']);
  expect(calculator(parsed.sheets)(parsed.sheets[0], 'E2')).toBe(212);
  expect(new Uint8Array(file.original!.data)).toEqual(new Uint8Array(bytes));
  const backup = await importFile({
    name: 'Saved.noffice',
    size: 1,
    text: async () => nativeBackup({ ...file, content: next }),
  } as File);
  expect(backup.content).toEqual(next);
});
it('updates imported calculated columns, preserves exceptions and repairs the master formula on rename', async () => {
  const { file, content } = await fixture(tableRenameFixture());
  const next = entry(content, 'C3', '=Sales[[#This Row],[Quantity]]*20');
  expect(calculator(next.sheets)(next.sheets[0], 'H4')).toBe(210);
  const renamed = await renameWorkbookTable(
    { ...file, content: next },
    {
      sheetId: next.sheets[0].id,
      table: 'Sales',
      name: 'Sales',
      column: { index: 1, name: 'Units' },
    },
  );
  expect(renamed.sheets[0].tables![0].calculatedColumns![2]).toBe('Sales[[#This Row],[Units]]*20');
  expect(entry(renamed, 'C2', '=Sales[[#This Row],[Units]]*30').sheets[0].cells.C4.value).toBe(
    '=Sales[[#This Row],[Units]]*30',
  );
});
it('restores missing master metadata from a legacy snapshot without changing edits or revision', async () => {
  const { file, content } = await fixture(tableRenameFixture());
  const original = content.sheets[0].tables![0].calculatedColumns;
  delete content.sheets[0].tables![0].calculatedColumns;
  content.sheets[0].cells.C4 = { value: '99' };
  const hydrated = await hydrateTableMetadata(file);
  expect(hydrated.revision).toBe(file.revision);
  const next = hydrated.content as WorkbookContent;
  expect(next.sheets[0].tables![0].calculatedColumns).toEqual(original);
  expect(next.sheets[0].cells.C4.value).toBe('99');
  expect(entry(next, 'C3', '=B3*20')).toBeUndefined();
});
it('exports a new calculated table, keeps unchanged imported bytes and fills resized rows', async () => {
  const next = entry(fresh(), 'B3', '=A3*10');
  const output = await (await exportOffice({ ...newFile('excel'), content: next })).arrayBuffer();
  const imported = await importWorkbook(output);
  expect(values(imported)).toEqual([20, 30, 50, 100]);
  expect(imported.sheets[0].tables![0].calculatedColumns).toEqual([null, 'A2*10']);
  expect(
    editSheetTable(next, next.sheets[0].id, 'Sales', 'A1:B5', defaultTableStyle).sheets[0].cells.B5
      .value,
  ).toBe('=A5*10');
  const { file, bytes } = await fixture(tableRenameFixture());
  expect(new Uint8Array(await (await exportOffice(file)).arrayBuffer())).toEqual(
    new Uint8Array(bytes),
  );
});
it.each(['array-master', 'array-range', 'external-data'])(
  'retains %s metadata and rejects automatic replacement',
  async (kind) => {
    const zip = await JSZip.loadAsync(await tableRenameFixture());
    const path = 'xl/tables/table1.xml';
    if (kind === 'array-master')
      zip.file(
        path,
        (await zip.file(path)!.async('string')).replace(
          '<calculatedColumnFormula>',
          '<calculatedColumnFormula array="1">',
        ),
      );
    if (kind === 'external-data')
      zip.file(
        path,
        (await zip.file(path)!.async('string')).replace(
          '<table ',
          '<table tableType="queryTable" ',
        ),
      );
    if (kind === 'array-range') {
      const sheet = 'xl/worksheets/sheet1.xml';
      zip.file(
        sheet,
        (await zip.file(sheet)!.async('string')).replace(
          '<c r="C2"><f>',
          '<c r="C2"><f t="array" ref="C2:C4">',
        ),
      );
    }
    const { file, content, bytes } = await fixture(zip.generateAsync({ type: 'arraybuffer' }));
    expect(content.sheets[0].tables![0].calculationBlocked).toMatch(/extended formula/);
    expect(() => entry(content, 'C3', '=B3*20')).toThrow(/extended formula/);
    expect(new Uint8Array(await (await exportOffice(file)).arrayBuffer())).toEqual(
      new Uint8Array(bytes),
    );
  },
);
it('rejects overlong or control-character formulas without changing a table', () => {
  const content = fresh();
  for (const formula of ['=' + '1'.repeat(8193), '=1\u0001'])
    expect(() => entry(content, 'B2', formula)).toThrow(/8,192/);
  expect(content.sheets[0].tables![0].calculatedColumns).toEqual([null, null]);
});
it('rejects inconsistent calculated-column metadata at export', async () => {
  const { file, content } = await fixture();
  content.sheets[0].tables![0].calculatedColumns = ['1'];
  await expect(exportOffice(file)).rejects.toThrow(/metadata must match/);
});
it('hydrates a legacy table created after import when it has no original table part', async () => {
  const { file, content } = await fixture();
  const next = createSheetTable(content, content.sheets[0].id, 'F3:G6', 'Supplies');
  delete next.sheets[0].tables![1].calculatedColumns;
  const hydrated = await hydrateTableMetadata({ ...file, content: next });
  const result = hydrated.content as WorkbookContent;
  expect(result.sheets[0].tables![1].calculatedColumns).toEqual([null, null]);
  expect(result.sheets[0].cells).toEqual(next.sheets[0].cells);
  expect((await exportOffice(hydrated)).size).toBeGreaterThan(0);
});
