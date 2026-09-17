import { expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { newFile, type Sheet, type OfficeFile } from './model';
import { importWorkbook, parseXML, elements } from './xlsx-import';
import { exportRetainedWorkbook } from './xlsx-preserve';
import { freezeSheet, writeFrozenPanes } from './sheet-panes';
import { sheetLayout } from './sheet-layout';

const base: Sheet = {
  id: 's',
  name: 'Sheet',
  rows: 10000,
  cols: 20,
  cells: { A1: { value: 'Heading' }, B3: { value: 'Input' } },
};
const bytes = (blob: Blob) =>
  new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = reject;
    reader.readAsArrayBuffer(blob);
  });

it('freezes above/left of the active cell and supports row, column and unfreeze without changing data', () => {
  const frozen = freezeSheet(base, 'selection', 'C4');
  expect(frozen).toMatchObject({ frozenRows: 3, frozenColumns: 2 });
  expect(frozen.cells).toBe(base.cells);
  expect(freezeSheet(frozen, 'row', 'C4')).toMatchObject({ frozenRows: 1, frozenColumns: 0 });
  expect(freezeSheet(frozen, 'column', 'C4')).toMatchObject({ frozenRows: 0, frozenColumns: 1 });
  expect(freezeSheet(frozen, 'none', 'C4')).toMatchObject({ frozenRows: 0, frozenColumns: 0 });
  expect(() => freezeSheet(base, 'selection', 'A1')).toThrow(/Select a cell/);
  expect(() => freezeSheet({ ...base, merges: ['A1:C2'] }, 'selection', 'B2')).toThrow(/merged/);
  expect(() => freezeSheet(base, 'selection', 'U10001')).toThrow(/available to scroll/);
});
it('keeps only frozen and visible body rows mounted with hidden/custom rows and merge anchors', () => {
  const layout = sheetLayout({
    ...base,
    frozenRows: 3,
    hiddenRows: [1],
    rowHeights: { 0: 60, 2: 45 },
    merges: ['A200:B202'],
  });
  const window = layout.window(6000);
  expect(window.rows.slice(0, 2)).toEqual([0, 2]);
  expect(window.rows.length).toBeLessThan(60);
  expect(window.rows).toContain(199);
  expect(new Set(window.rows).size).toBe(window.rows.length);
  expect(window.top + window.bottom + window.rows.reduce((n, r) => n + layout.height(r), 0)).toBe(
    layout.offsets.at(-1),
  );
  const last = layout.window(9999999);
  expect(last.rows.at(-1)).toBe(9999);
  expect(last.bottom).toBe(0);
  const shortRows = sheetLayout({ ...base, defaultRowHeight: 8, frozenRows: 2 });
  const tallWindow = shortRows.window(1000, 1600);
  expect(shortRows.offsets[tallWindow.end]).toBeGreaterThanOrEqual(2600);
  expect(tallWindow.rows.length).toBeLessThan(215);
});
it('writes schema-ordered pane/selection elements and retains other view metadata/windows', () => {
  const doc = parseXML(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C4"/><sheetViews><sheetView workbookViewId="0" zoomScale="85" showGridLines="0"><pane state="split" xSplit="500"/><selection pane="topRight"/><extLst/></sheetView><sheetView workbookViewId="1" zoomScale="50"/></sheetViews><sheetData/></worksheet>',
  );
  const other = new XMLSerializer().serializeToString(elements(doc, 'sheetView')[1]);
  writeFrozenPanes(doc.documentElement, 3, 2);
  const view = elements(doc, 'sheetView')[0],
    pane = elements(doc, 'pane')[0];
  expect(pane.getAttribute('topLeftCell')).toBe('C4');
  expect(pane.getAttribute('activePane')).toBe('bottomRight');
  expect(Array.from(view.children).map((e) => e.localName)).toEqual([
    'pane',
    'selection',
    'extLst',
  ]);
  expect(view.getAttribute('zoomScale')).toBe('85');
  expect(view.getAttribute('showGridLines')).toBe('0');
  expect(new XMLSerializer().serializeToString(elements(doc, 'sheetView')[1])).toBe(other);
  writeFrozenPanes(doc.documentElement, 0, 0);
  expect(elements(doc, 'pane')).toHaveLength(0);
  expect(elements(doc, 'selection')[0].hasAttribute('pane')).toBe(false);
  const missing = parseXML(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr/><dimension/><sheetFormatPr/><sheetData/></worksheet>',
  );
  writeFrozenPanes(missing.documentElement, 1, 0);
  expect(Array.from(missing.documentElement.children).map((e) => e.localName)).toEqual([
    'sheetPr',
    'dimension',
    'sheetViews',
    'sheetFormatPr',
    'sheetData',
  ]);
  expect(() => writeFrozenPanes(missing.documentElement, NaN, 0)).toThrow();
});
it.each([
  [3, 2],
  [1, 0],
  [0, 1],
  [0, 0],
])(
  'retains source parts and native view settings after exporting %i rows/%i columns',
  async (rows, columns) => {
    const workbook = new ExcelJS.Workbook(),
      sheet = workbook.addWorksheet('Source');
    sheet.getCell('A1').value = 'Heading';
    sheet.getCell('C4').value = { formula: '1+2', result: 3 };
    sheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 2, zoomScale: 85, showGridLines: false }];
    sheet.pageSetup = { orientation: 'landscape' };
    const original = (await workbook.xlsx.writeBuffer()) as ArrayBuffer;
    const content = await importWorkbook(original);
    content.sheets[0] = { ...content.sheets[0], frozenRows: rows, frozenColumns: columns };
    const file: OfficeFile = {
      ...newFile('excel', 'Source'),
      content,
      original: {
        name: 'Source.xlsx',
        data: original,
      },
    };
    const output = await bytes(await exportRetainedWorkbook(file));
    const reimport = await importWorkbook(output);
    expect(reimport.sheets[0]).toMatchObject({
      frozenRows: rows,
      frozenColumns: columns,
      gridLines: false,
    });
    const before = await JSZip.loadAsync(original),
      after = await JSZip.loadAsync(output);
    for (const path of Object.keys(before.files))
      if (!before.files[path].dir && path !== content.sheets[0].sourcePath)
        expect(await after.file(path)!.async('uint8array'), path).toEqual(
          await before.file(path)!.async('uint8array'),
        );
    const doc = parseXML(await after.file(content.sheets[0].sourcePath!)!.async('string'));
    expect(elements(doc, 'sheetView')[0].getAttribute('zoomScale')).toBe('85');
    expect(elements(doc, 'f')[0].textContent).toBe('1+2');
  },
);
