import { it, expect, beforeAll, afterAll } from 'vitest';
import { newFile, type WorkbookContent } from './model';
import { createSheetTable } from './sheet-tables';
import { calculatedColumnEntry } from './sheet-calculated-columns';
import { tableRowEntry, syncTableTotals } from './sheet-table-entry';
import { calculator } from './formulas';
import { exportOffice } from './formats';
import { importWorkbook } from './xlsx-import';
import { tableEntryFixture } from '../tests/fixtures/xlsx-table-entry';
import { importFile } from './formats';
import { hydrateTableMetadata } from './workbook-tables';
import JSZip from 'jszip';
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
  const content = newFile('excel').content as WorkbookContent,
    s = content.sheets[0];
  s.cells = {
    A1: { value: 'Quantity' },
    B1: { value: 'Total' },
    A2: { value: '2' },
    A3: { value: '3' },
    D1: { value: '=SUM(Sales[Total])' },
  };
  const table = createSheetTable(content, s.id, 'A1:B3', 'Sales');
  return calculatedColumnEntry(table, s.id, 'B2', '=A2*10')!;
}
const enter = (content: WorkbookContent, ref: string, value: string) =>
  tableRowEntry(content, content.sheets[0].id, ref, value);
it('retains several pending row additions and totals edits across export and legacy hydration', async () => {
  const input = await tableEntryFixture(),
    file = await importFile(new File([input], 'Entry.xlsx'));
  let content = file.content as WorkbookContent;
  delete content.sheets[0].tables![0].totalLabels;
  delete content.sheets[0].tables![0].totalFormulas;
  content = (await hydrateTableMetadata(file)).content as WorkbookContent;
  expect(content.sheets[0].tables![0].totalLabels![0]).toBe('Total');
  content = enter(content, 'B5', '7')!;
  content = enter(content, 'B6', '11')!;
  content = enter(content, 'B7', '=2+3')!;
  const next = structuredClone(content);
  next.sheets[0].cells.C7 = { value: '99' };
  content = syncTableTotals(content, next);
  const output = await (await exportOffice({ ...file, content })).arrayBuffer(),
    imported = await importWorkbook(output);
  expect(imported.sheets[0].tables![0]).toMatchObject({
    ref: 'A1:C7',
    totalRows: 1,
    totalFormulas: [null, '2+3', null],
    totalLabels: ['Total', null, '99'],
  });
  expect(calculator(imported.sheets)(imported.sheets[0], 'H4')).toBe(308);
  expect(imported.sheets[0].cells.C7.dataType).toBe('text');
  const zip = await JSZip.loadAsync(output),
    original = await JSZip.loadAsync(input);
  expect(await zip.file('custom/preservation.xml')!.async('string')).toBe(
    await original.file('custom/preservation.xml')!.async('string'),
  );
}, 20000);
it('retains preset totals metadata while refusing implicit activation', async () => {
  const zip = await JSZip.loadAsync(await tableEntryFixture());
  zip.file(
    'xl/tables/table1.xml',
    (await zip.file('xl/tables/table1.xml')!.async('string')).replace(
      'totalsRowFunction="none"',
      'totalsRowFunction="sum"',
    ),
  );
  const content = await importWorkbook(await zip.generateAsync({ type: 'arraybuffer' }));
  expect(content.sheets[0].tables![0].totalsActivationBlocked).toContain('preset');
  expect(enter(content, 'B5', '=2+3')).toBeUndefined();
}, 20000);
it.each(['7', '0'])(
  'expands from literal %s and fills calculated cells in one immutable change',
  (value) => {
    const before = fresh(),
      next = enter(before, 'A4', value)!;
    expect(next.sheets[0].tables![0].ref).toBe('A1:B4');
    expect(next.sheets[0].autoFilters![0].ref).toBe('A1:B4');
    expect(next.sheets[0].cells.B4.value).toBe('=A4*10');
    expect(calculator(next.sheets)(next.sheets[0], 'D1')).toBe(50 + Number(value) * 10);
    expect(before.sheets[0].cells.A4).toBeUndefined();
  },
);
it('preserves an entered calculated-column exception without changing existing rows or the master', () => {
  const next = enter(fresh(), 'B4', '99')!;
  expect(next.sheets[0].cells.B4.value).toBe('99');
  expect(next.sheets[0].cells.B3.value).toBe('=A3*10');
  expect(next.sheets[0].tables![0].calculatedColumns![1]).toBe('A2*10');
});
it.each([
  ['A4', '=4+5'],
  ['B4', '=A4*20'],
])('activates a totals row on formula entry at %s, excluded from data sums', (ref, value) => {
  const next = enter(fresh(), ref, value)!;
  expect(next.sheets[0].tables![0].ref).toBe('A1:B4');
  expect(next.sheets[0].cells[ref].value).toBe(value);
  expect(next.sheets[0].tables![0].totalRows).toBe(1);
  expect(calculator(next.sheets)(next.sheets[0], 'D1')).toBe(50);
  if (ref === 'A4') expect(next.sheets[0].cells.B4).toBeUndefined();
});
it('restores retained totals labels and formulas when totals become visible', () => {
  const content = fresh(),
    t = content.sheets[0].tables![0];
  t.totalLabels = ['Total', null];
  const next = enter(content, 'B4', '=SUM(Sales[Total])')!;
  expect(next.sheets[0].cells.A4.value).toBe('Total');
  expect(calculator(next.sheets)(next.sheets[0], 'B4')).toBe(50);
  expect(next.sheets[0].tables![0].totalFormulas).toEqual([null, 'SUM(Sales[Total])']);
});
it('synchronizes totals formulas, literal labels and deletion without changing body masters', () => {
  let content = enter(fresh(), 'B4', '=SUM(Sales[Total])')!;
  for (const value of ['99', "'=1", '=2+3', '']) {
    const next = structuredClone(content);
    next.sheets[0].cells.B4 = { value };
    content = syncTableTotals(content, next);
    const table = content.sheets[0].tables![0];
    expect(table.totalFormulas![1]).toBe(value.startsWith('=') ? value.slice(1) : null);
    expect(table.totalLabels![1]).toBe(
      value.startsWith('=') ? null : value.replace(/^'/, '') || null,
    );
    expect(table.calculatedColumns![1]).toBe('A2*10');
    expect(calculator(content.sheets)(content.sheets[0], 'D1')).toBe(50);
  }
});
it('leaves unsupported totals activation as ordinary entry and validates restored totals atomically', () => {
  const content = fresh(),
    s = content.sheets[0];
  s.tables![0].totalsActivationBlocked = 'Preset totals';
  expect(enter(content, 'A4', '=1')).toBeUndefined();
  expect(enter(content, 'A4', '7')!.sheets[0].tables![0].totalRows).toBe(0);
  delete s.tables![0].totalsActivationBlocked;
  s.tables![0].totalFormulas = [null, '99'];
  s.cells.B4 = {
    value: '',
    validation: {
      type: 'whole',
      operator: 'lessThan',
      formulae: [5],
      showErrorMessage: true,
      error: 'Too large',
    },
  };
  expect(() => enter(content, 'A4', '=1')).toThrow('Too large');
  expect(s.tables![0].totalRows).toBe(0);
});
it('exports and reimports new-workbook totals metadata and excludes totals from structured sums', async () => {
  const content = fresh();
  content.sheets[0].tables![0].totalLabels = ['Total', null];
  const next = enter(content, 'B4', '=SUM(Sales[Total])')!;
  const imported = await importWorkbook(
    await (await exportOffice({ ...newFile('excel'), content: next })).arrayBuffer(),
  );
  expect(imported.sheets[0].tables![0]).toMatchObject({
    totalRows: 1,
    totalLabels: ['Total', null],
    totalFormulas: [null, 'SUM(Sales[Total])'],
  });
  expect(calculator(imported.sheets)(imported.sheets[0], 'B4')).toBe(50);
  expect(calculator(imported.sheets)(imported.sheets[0], 'D1')).toBe(50);
}, 20000);
it('can expand while replacing an existing value in the entered cell', () => {
  const content = fresh();
  content.sheets[0].cells.A4 = { value: '7' };
  expect(enter(content, 'A4', '7')!.sheets[0].cells.B4.value).toBe('=A4*10');
});
it.each(['99', '=1', '=""'])('leaves expansion off when another cell contains %s', (value) => {
  const content = fresh();
  content.sheets[0].cells.B4 = { value };
  expect(enter(content, 'A4', '7')).toBeUndefined();
});
it.each([
  ['A5', '7'],
  ['C4', '7'],
  ['A4', ''],
  ['A3', '7'],
])('does not expand for entry at %s with %j', (ref, value) => {
  expect(enter(fresh(), ref, value)).toBeUndefined();
});
it('keeps unsupported structures, overlapping tables and protected sheets on the ordinary entry path', () => {
  for (const kind of ['protected', 'totals', 'merged', 'filter', 'blocked', 'limit', 'overlap']) {
    const content = fresh(),
      s = content.sheets[0];
    if (kind === 'protected') s.protected = true;
    if (kind === 'totals') s.tables![0].totalRows = 1;
    if (kind === 'merged') s.merges = ['A4:B4'];
    if (kind === 'filter') s.autoFilters![0].columns = [{ col: 0, values: ['2'] }];
    if (kind === 'blocked') s.tables![0].resizeBlocked = 'Extended table';
    if (kind === 'overlap') s.tables!.push({ ...s.tables![0], name: 'Other', ref: 'A4:B5' });
    if (kind === 'limit') s.tables![0].ref = 'A1:B10000';
    expect(enter(content, kind === 'limit' ? 'A10001' : 'A4', '7')).toBeUndefined();
  }
});
it('rejects invalid entry or generated cells atomically', () => {
  for (const ref of ['A4', 'B4']) {
    const content = fresh();
    content.sheets[0].cells[ref] = {
      value: '',
      validation: {
        type: 'whole',
        operator: 'lessThan',
        formulae: [5],
        showErrorMessage: true,
        error: 'Too large',
      },
    };
    expect(() => enter(content, 'A4', '7')).toThrow('Too large');
    expect(content.sheets[0].tables![0].ref).toBe('A1:B3');
  }
});
it('preserves note/style metadata and exports the extended table and master', async () => {
  const content = fresh();
  content.sheets[0].cells.B4 = { value: '', numFmt: '0.00', note: 'Keep' };
  const next = enter(content, 'A4', '7')!;
  expect(next.sheets[0].cells.B4).toMatchObject({ numFmt: '0.00', note: 'Keep' });
  const imported = await importWorkbook(
    await (await exportOffice({ ...newFile('excel'), content: next })).arrayBuffer(),
  );
  expect(imported.sheets[0].tables![0].ref).toBe('A1:B4');
  expect(imported.sheets[0].tables![0].calculatedColumns![1]).toBe('A2*10');
  expect(calculator(imported.sheets)(imported.sheets[0], 'D1')).toBe(120);
});
