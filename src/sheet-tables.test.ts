import { it, expect, beforeAll, afterAll } from 'vitest';
import JSZip from 'jszip';
import { newFile, type WorkbookContent } from './model';
import {
  createSheetTable,
  editSheetTable,
  tableBounds,
  tableAt,
  defaultTableStyle,
  tableCellStyle,
  tableHeaderError,
} from './sheet-tables';
import { calculator } from './formulas';
import { tableFixture } from '../tests/fixtures/xlsx-tables';
import { importFile, exportOffice, nativeBackup } from './formats';
import { importWorkbook, elements, parseXML } from './xlsx-import';
import { readAutoFilters, writeFilterColumns, applySheetFilters } from './sheet-filters';
import { hydrateTableMetadata } from './workbook-tables';
const nativeArrayBuffer = Blob.prototype.arrayBuffer;
beforeAll(() => {
  if (!nativeArrayBuffer)
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
  if (!nativeArrayBuffer) delete (Blob.prototype as Partial<Blob>).arrayBuffer;
});
function content() {
  const file = newFile('excel');
  const value = file.content as WorkbookContent;
  value.sheets[0].cells = {
    A1: { value: 'Amount' },
    B1: { value: 'Amount' },
    A2: { value: '2' },
    B2: { value: '3' },
    D1: { value: '=SUM(Orders[Amount])' },
    D2: { value: '=D1*2' },
  };
  return { file, value, sheet: value.sheets[0] };
}
it('creates a named table, normalizes headers and resolves structured references without mutating input', () => {
  const { value, sheet } = content();
  const next = createSheetTable(value, sheet.id, '$A$1:$C$3', 'Orders');
  expect(next.sheets[0].tables![0].columns).toEqual(['Amount', 'Amount2', 'Column3']);
  expect(sheet.tables).toBeUndefined();
  expect(sheet.cells.C1).toBeUndefined();
  expect(calculator(next.sheets)(next.sheets[0], 'D1')).toBe(2);
  expect(tableAt(next.sheets[0], 'B2')?.name).toBe('Orders');
  expect(tableAt(next.sheets[0], 'D3')).toBeUndefined();
});
it.each(['A1', 'R', 'C', 'R1C1', 'bad name', '1Table', 'Orders!', 'XFD1048576'])(
  'rejects invalid table name %s',
  (name) => {
    const { value, sheet } = content();
    expect(() => createSheetTable(value, sheet.id, 'A1:B2', name)).toThrow();
  },
);
it.each(['A1', 'B3:A1', 'A0:B2', 'A1:IW2', 'A1:A10001', 'A1:IV10000'])(
  'rejects invalid or overlarge table range %s',
  (ref) => {
    const { value, sheet } = content();
    expect(() => createSheetTable(value, sheet.id, ref, 'Orders')).toThrow();
  },
);
it('rejects protection, name collisions, merges and filter/table overlaps', () => {
  const { value, sheet } = content();
  sheet.protected = true;
  expect(() => createSheetTable(value, sheet.id, 'A1:B2', 'Orders')).toThrow(/protected/);
  delete sheet.protected;
  sheet.merges = ['A1:B1'];
  expect(() => createSheetTable(value, sheet.id, 'A1:B2', 'Orders')).toThrow(/Unmerge/);
  delete sheet.merges;
  sheet.autoFilters = [{ ref: 'A1:B3', columns: [] }];
  expect(() => createSheetTable(value, sheet.id, 'A1:B2', 'Orders')).toThrow(/filter/);
  delete sheet.autoFilters;
  const next = createSheetTable(value, sheet.id, 'A1:B2', 'Orders');
  expect(() => createSheetTable(next, sheet.id, 'F1:G2', 'orders')).toThrow(/already/);
  expect(() => createSheetTable(next, sheet.id, 'B1:C3', 'Second')).toThrow(/overlap/);
});
it('resizes rows, preserves excluded values and recalculates structured dependencies', () => {
  const { value, sheet } = content();
  sheet.cells.A3 = { value: '7' };
  const created = createSheetTable(value, sheet.id, 'A1:B2', 'Orders');
  const next = editSheetTable(created, sheet.id, 'Orders', 'A1:B3', {
    ...defaultTableStyle,
    firstColumn: true,
  });
  expect(calculator(next.sheets)(next.sheets[0], 'D2')).toBe(18);
  const shrunk = editSheetTable(next, sheet.id, 'Orders', 'A1:B2', defaultTableStyle);
  expect(shrunk.sheets[0].cells.A3.value).toBe('7');
  expect(calculator(shrunk.sheets)(shrunk.sheets[0], 'D2')).toBe(4);
  expect(shrunk.sheets[0].autoFilters![0].ref).toBe('A1:B2');
  expect(() => editSheetTable(next, sheet.id, 'Orders', 'A1:C3', defaultTableStyle)).toThrow(
    /columns/,
  );
});
it('guards header edits, active filters, calculated columns and totals', () => {
  const { value, sheet } = content();
  const created = createSheetTable(value, sheet.id, 'A1:B2', 'Orders'),
    s = created.sheets[0];
  expect(tableHeaderError(s, { ...s, cells: { ...s.cells, A1: { value: 'Renamed' } } })).toMatch(
    /structured references/,
  );
  s.autoFilters![0].columns = [{ col: 0, values: ['2'] }];
  expect(() => editSheetTable(created, s.id, 'Orders', 'A1:B3', defaultTableStyle)).toThrow(
    /criteria/,
  );
  s.autoFilters![0].columns = [];
  s.tables![0].totalRows = 1;
  expect(() => editSheetTable(created, s.id, 'Orders', 'A1:B3', defaultTableStyle)).toThrow(
    /totals/,
  );
  s.tables![0].totalRows = 0;
  s.tables![0].resizeBlocked = 'Calculated columns';
  expect(() => editSheetTable(created, s.id, 'Orders', 'A1:B3', defaultTableStyle)).toThrow(
    /Calculated/,
  );
});
it('renders supported theme styles without overwriting direct cell formats', () => {
  const { value, sheet } = content();
  const s = createSheetTable(value, sheet.id, 'A1:B3', 'Orders').sheets[0];
  expect(tableCellStyle(s, 'A1')).toMatchObject({ fill: '#4F81BD', color: '#ffffff', bold: true });
  expect(tableCellStyle(s, 'A2').fill).toBe('#dce6f1');
  expect(tableCellStyle(s, 'A3').fill).toBe('#ffffff');
  expect(tableBounds('a1:$b$3').ref).toBe('A1:B3');
});
it('exports new tables with valid parts, names, filters and formulas, then survives native backup', async () => {
  const { value, sheet, file } = content();
  file.content = createSheetTable(value, sheet.id, 'A1:C3', 'Orders');
  const backup = nativeBackup(file);
  const restored = await importFile({
    name: 'tables.noffice',
    size: backup.length,
    text: async () => backup,
  } as File);
  const data = await (await exportOffice(restored)).arrayBuffer(),
    imported = await importWorkbook(data);
  expect(imported.sheets[0].tables![0]).toMatchObject({
    name: 'Orders',
    ref: 'A1:C3',
    columns: ['Amount', 'Amount2', 'Column3'],
    style: defaultTableStyle,
  });
  expect(calculator(imported.sheets)(imported.sheets[0], 'D2')).toBe(4);
  expect(imported.sheets[0].autoFilters).toHaveLength(1);
  const zip = await JSZip.loadAsync(data);
  expect(await zip.file('xl/workbook.xml')!.async('string')).not.toContain('defaultThemeVersion');
});
it('keeps large imported table extents readable while bounding creation and resizing', () => {
  const { sheet } = content();
  sheet.tables = [
    {
      name: 'Huge',
      ref: 'A1:IV10000',
      columns: [],
      headerRows: 1,
      totalRows: 0,
      style: defaultTableStyle,
    },
  ];
  expect(tableAt(sheet, 'D9000')?.name).toBe('Huge');
  expect(tableCellStyle(sheet, 'D9000').fill).toBeDefined();
});
it('keeps table dropdown visibility separate from criteria and preserves button attributes when filtering', () => {
  const doc = parseXML(
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1:B4"><filterColumn colId="0" hiddenButton="1"/><filterColumn colId="1" showButton="0"/></autoFilter></table>',
  );
  expect(readAutoFilters(doc.documentElement)[0].columns).toEqual([]);
  const filter = elements(doc, 'autoFilter')[0];
  writeFilterColumns(filter, [{ col: 0, values: ['2'] }]);
  expect(
    elements(filter, 'filterColumn')
      .find((c) => c.getAttribute('colId') === '0')
      ?.getAttribute('hiddenButton'),
  ).toBe('1');
  expect(readAutoFilters(doc.documentElement)[0].columns).toEqual([
    { col: 0, values: ['2'], blank: false },
  ]);
  writeFilterColumns(filter, []);
  expect(readAutoFilters(doc.documentElement)[0].columns).toEqual([]);
  expect(elements(filter, 'filterColumn')).toHaveLength(2);
});
it('exports criteria on a newly created table and retains hidden filtered rows', async () => {
  const { value, sheet, file } = content();
  sheet.cells.A3 = { value: '7' };
  const next = createSheetTable(value, sheet.id, 'A1:B3', 'Orders'),
    s = next.sheets[0];
  next.sheets[0] = applySheetFilters(s, next.sheets, [
    { ...s.autoFilters![0], columns: [{ col: 0, values: ['7'] }] },
  ]);
  file.content = next;
  const imported = await importWorkbook(await (await exportOffice(file)).arrayBuffer());
  expect(imported.sheets[0].hiddenRows).toContain(1);
  expect(imported.sheets[0].autoFilters![0].columns[0].values).toEqual(['7']);
});
it('patches original table XML, structured formula caches and unrelated parts, including consecutive exports', async () => {
  const input = await tableFixture(),
    file = await importFile(new File([input], 'source.xlsx'));
  let value = file.content as WorkbookContent;
  const id = value.sheets[0].id;
  expect(new Uint8Array(await (await exportOffice(file)).arrayBuffer())).toEqual(
    new Uint8Array(input),
  );
  value = editSheetTable(value, id, 'Sales', 'A1:C6', {
    ...defaultTableStyle,
    name: 'TableStyleMedium4',
    firstColumn: true,
  });
  file.content = createSheetTable(value, id, 'F3:G6', 'Supplies');
  const bytes = await (await exportOffice(file)).arrayBuffer(),
    zip = await JSZip.loadAsync(bytes),
    original = await JSZip.loadAsync(input);
  const next = await importWorkbook(bytes);
  expect(next.sheets[0].tables).toHaveLength(2);
  expect(next.sheets[0].cells.E1.cachedValue).toBe(28);
  expect(next.sheets[0].cells.E2.cachedValue).toBe(56);
  for (const path of [
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/worksheets/sheet2.xml',
    'custom/preservation.xml',
    'xl/comments1.xml',
  ])
    expect(await zip.file(path)!.async('string')).toBe(await original.file(path)!.async('string'));
  const table = parseXML(await zip.file(next.sheets[0].tables![0].sourcePath!)!.async('string'));
  expect(table.documentElement.getAttribute('ref')).toBe('A1:C6');
  expect(elements(table, 'autoFilter')[0].getAttribute('ref')).toBe('A1:C6');
  const imported = await importFile(new File([bytes], 'edited.xlsx'));
  imported.content = editSheetTable(
    imported.content as WorkbookContent,
    (imported.content as WorkbookContent).sheets[0].id,
    'Sales',
    'A1:C4',
    defaultTableStyle,
  );
  expect(
    (await importWorkbook(await (await exportOffice(imported)).arrayBuffer())).sheets[0].cells.E1
      .cachedValue,
  ).toBe(10);
});
it('restores legacy table metadata for existing workbooks without losing edits or unchanged bytes', async () => {
  const input = await tableFixture(),
    file = await importFile(new File([input], 'legacy.xlsx'));
  const content = file.content as WorkbookContent;
  for (const s of content.sheets) {
    delete s.tableTheme;
    for (const t of s.tables || []) {
      delete t.sourcePath;
      delete t.style;
      delete t.resizeBlocked;
    }
  }
  expect(new Uint8Array(await (await exportOffice(file)).arrayBuffer())).toEqual(
    new Uint8Array(input),
  );
  content.sheets[0].cells.B2 = { value: '15' };
  const upgraded = await hydrateTableMetadata(file),
    next = upgraded.content as WorkbookContent;
  expect(next.sheets[0].tables![0].sourcePath).toBe('xl/tables/table1.xml');
  expect(next.sheets[0].cells.B2.value).toBe('15');
  expect(upgraded.revision).toBe(file.revision);
  expect(content.sheets[0].tables![0].sourcePath).toBeUndefined();
  const exported = await importWorkbook(await (await exportOffice(upgraded)).arrayBuffer());
  expect(exported.sheets[0].cells.E1.cachedValue).toBe(23);
});
