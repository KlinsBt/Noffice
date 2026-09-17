import { it, expect, beforeAll, afterAll } from 'vitest';
import JSZip from 'jszip';
import { newFile, type WorkbookContent } from './model';
import {
  createSheetTable,
  editSheetTable,
  defaultTableStyle,
  legacyCalculatedResizeBlock,
} from './sheet-tables';
import { calculatedColumnEntry, clearRemovedColumnFormulas } from './sheet-calculated-columns';
import { qualifyTableReferences } from './table-references';
import { calculator } from './formulas';
import { exportOffice, importFile } from './formats';
import { importWorkbook } from './xlsx-import';
import { hydrateTableMetadata } from './workbook-tables';
import { tableResizeFixture } from '../tests/fixtures/xlsx-table-resize';
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
function fresh(formula = '=Sales[[#This Row],[Quantity]]*10') {
  const content = newFile('excel').content as WorkbookContent,
    s = content.sheets[0];
  s.cells = {
    A1: { value: 'Quantity' },
    B1: { value: 'Total' },
    A2: { value: '2' },
    A3: { value: '3' },
    A4: { value: '5' },
    A5: { value: '7' },
    A6: { value: '11' },
    D1: { value: '=SUM(Sales[Total])' },
  };
  const table = createSheetTable(content, s.id, 'A1:B4', 'Sales');
  return calculatedColumnEntry(table, s.id, 'B2', formula)!;
}
const resize = (content: WorkbookContent, ref: string) =>
  editSheetTable(content, content.sheets[0].id, 'Sales', ref, defaultTableStyle);
const calc = (content: WorkbookContent, ref: string) =>
  calculator(content.sheets)(content.sheets[0], ref);
it('fills blank added rows from the master while preserving existing values, formulas, styles and exceptions', () => {
  const content = fresh();
  content.sheets[0].cells.B3.value = '99';
  content.sheets[0].cells.B5 = { value: '40' };
  content.sheets[0].cells.B6 = { value: '', numFmt: '0.00', note: 'Keep', bold: true };
  const next = resize(content, 'A1:B6');
  expect(['B2', 'B3', 'B4', 'B5', 'B6', 'D1'].map((ref) => calc(next, ref))).toEqual([
    20, 99, 50, 40, 110, 319,
  ]);
  expect(next.sheets[0].cells.B6).toMatchObject({ numFmt: '0.00', note: 'Keep', bold: true });
  expect(next.sheets[0].autoFilters![0].ref).toBe('A1:B6');
  expect(content.sheets[0].cells.B6.value).toBe('');
  content.sheets[0].cells.B5 = { value: '=A5*2' };
  expect(resize(content, 'A1:B6').sheets[0].cells.B5.value).toBe('=A5*2');
});
it.each(['99', ''])(
  'uses the stored master even when the final existing row has an exception %j',
  (value) => {
    const content = fresh();
    content.sheets[0].cells.B4.value = value;
    expect(calc(resize(content, 'A1:B6'), 'B6')).toBe(110);
  },
);
it('anchors relative references at the first data row and retains locks when extending', () => {
  const content = fresh('=A2*10+$A$2');
  const next = resize(content, 'A1:B6');
  expect(next.sheets[0].cells.B6.value).toBe('=A6*10+$A$2');
  expect(calc(next, 'B6')).toBe(112);
});
it('shrinks without deleting cells and binds excluded local selectors to the old table', () => {
  const content = fresh('=[@Quantity]*10');
  const next = resize(content, 'A1:B2');
  expect(next.sheets[0].cells.B3.value).toBe('=Sales[@Quantity]*10');
  expect(calc(next, 'B3')).toBe('#VALUE!');
  expect(calc(next, 'D1')).toBe(20);
  expect(content.sheets[0].cells.B3.value).toBe('=[@Quantity]*10');
  expect(calc(resize(next, 'A1:B6'), 'B3')).toBe(30);
  expect(calc(resize(next, 'A1:B6'), 'B6')).toBe(110);
});
it('keeps A1 formulas meaningful after shrinking', () => {
  const next = resize(fresh('=A2*10+$A$2'), 'A1:B2');
  expect(calc(next, 'B4')).toBe(52);
});
it('qualifies selectors without touching strings, names, sheet names or already qualified tables', () => {
  expect(
    qualifyTableReferences('=SUM([Quantity])+Other[Quantity]+"[Quantity]"+\'Sheet 1\'!A1', 'Sales'),
  ).toBe('=SUM(Sales[Quantity])+Other[Quantity]+"[Quantity]"+\'Sheet 1\'!A1');
  expect(qualifyTableReferences("=[@[Net '[USD']]]", 'Sales')).toBe("=Sales[@[Net '[USD']]]");
  expect(() => qualifyTableReferences('=[Book.xlsx]Sheet1!A1', 'Sales')).toThrow(/External/);
});
it('rejects an excluded external-reference formula without partially shrinking', () => {
  const content = fresh();
  content.sheets[0].cells.B4.value = '=[Book.xlsx]Sheet1!A1';
  expect(() => resize(content, 'A1:B2')).toThrow(/External/);
  expect(content.sheets[0].tables![0].ref).toBe('A1:B4');
});
it('clears the master only when an edit removes the last nonblank body value', () => {
  const before = fresh(),
    proposed = structuredClone(before);
  for (const ref of ['B2', 'B3', 'B4']) proposed.sheets[0].cells[ref] = { value: '' };
  const cleared = clearRemovedColumnFormulas(before, proposed);
  expect(cleared.sheets[0].tables![0].calculatedColumns).toEqual([null, null]);
  expect(resize(cleared, 'A1:B6').sheets[0].cells.B6).toBeUndefined();
  expect(before.sheets[0].tables![0].calculatedColumns![1]).toBeTruthy();
  proposed.sheets[0].cells.B3.value = '0';
  expect(
    clearRemovedColumnFormulas(before, proposed).sheets[0].tables![0].calculatedColumns![1],
  ).toBeTruthy();
});
it('rejects protected, overlapping, filtered and unsupported table resizes atomically', () => {
  for (const kind of ['protected', 'merged', 'filter', 'extended', 'limit', 'totals', 'columns']) {
    const content = fresh(),
      s = content.sheets[0];
    if (kind === 'protected') s.protected = true;
    if (kind === 'merged') s.merges = ['B5:C5'];
    if (kind === 'filter') s.autoFilters![0].columns = [{ col: 0, values: ['2'] }];
    if (kind === 'extended') s.tables![0].calculationBlocked = 'Unsupported formula';
    if (kind === 'totals') s.tables![0].totalRows = 1;
    const ref = kind === 'columns' ? 'A1:C6' : kind === 'limit' ? 'A1:B10001' : 'A1:B6';
    expect(() => resize(content, ref)).toThrow();
    expect(s.cells.B6).toBeUndefined();
  }
});
it('validates generated cells against the complete proposed workbook before committing', () => {
  const content = fresh();
  content.sheets[0].cells.B6 = {
    value: '',
    validation: {
      type: 'whole',
      operator: 'lessThan',
      formulae: [100],
      showErrorMessage: true,
      error: 'Too large',
    },
  };
  expect(() => resize(content, 'A1:B6')).toThrow('B6: Too large');
  expect(content.sheets[0].cells.B5).toBeUndefined();
});
async function fixture() {
  const bytes = await tableResizeFixture();
  const file = await importFile({
    name: 'Resize.xlsx',
    size: bytes.byteLength,
    arrayBuffer: async () => bytes,
  } as File);
  if (file.content.kind !== 'excel') throw Error('Workbook');
  return { file, content: file.content, bytes };
}
it('retains original package parts while updating table/filter ranges, row formulas and caches', async () => {
  const { file, content, bytes } = await fixture();
  expect(content.sheets[0].tables![0].resizeBlocked).toBeUndefined();
  const next = resize(content, 'A1:C6'),
    output = await (await exportOffice({ ...file, content: next })).arrayBuffer();
  const zip = await JSZip.loadAsync(output),
    original = await JSZip.loadAsync(bytes);
  const xml = await zip.file('xl/tables/table1.xml')!.async('string');
  expect(xml).toContain('ref="A1:C6"');
  expect(xml).toContain(
    '<calculatedColumnFormula>Sales[[#This Row],[Quantity]]*10</calculatedColumnFormula>',
  );
  const reimport = await importWorkbook(output);
  expect(calc(reimport, 'C6')).toBe(110);
  expect(calc(reimport, 'H4')).toBe(278);
  for (const path of [
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/comments1.xml',
    'xl/worksheets/sheet2.xml',
    'custom/preservation.xml',
  ])
    expect(await zip.file(path)!.async('string')).toBe(await original.file(path)!.async('string'));
  expect(new Uint8Array(file.original!.data)).toEqual(new Uint8Array(bytes));
});
it('migrates the former calculated-column resize restriction without removing real source guards', async () => {
  const { file, content } = await fixture();
  content.sheets[0].tables![0].resizeBlocked = legacyCalculatedResizeBlock;
  const hydrated = await hydrateTableMetadata(file);
  expect(hydrated.revision).toBe(file.revision);
  expect((hydrated.content as WorkbookContent).sheets[0].tables![0].resizeBlocked).toBeUndefined();
  expect(resize(hydrated.content as WorkbookContent, 'A1:C6').sheets[0].cells.C6.value).toContain(
    'Quantity',
  );
});
it('exports resized new tables and preserves an explicit empty-string text destination', async () => {
  const content = fresh('=A2*10');
  content.sheets[0].cells.B5 = { value: '', dataType: 'text' };
  const next = resize(content, 'A1:B6');
  expect(next.sheets[0].cells.B5.dataType).toBe('text');
  const parsed = await importWorkbook(
    await (await exportOffice({ ...newFile('excel'), content: next })).arrayBuffer(),
  );
  expect(calc(parsed, 'B6')).toBe(110);
  expect(parsed.sheets[0].tables![0].ref).toBe('A1:B6');
});
