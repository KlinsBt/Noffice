import { expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { applySheetFilters } from './sheet-filters';
import { importWorkbook, parseXML, elements } from './xlsx-import';
import { exportOffice } from './formats';
import { newFile, contentSchema, type Sheet } from './model';
async function bytes(blob: Blob) {
  return new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = reject;
    reader.readAsArrayBuffer(blob);
  });
}
for (const kind of ['new', 'worksheet', 'table'] as const)
  it(`${kind} filter edits export criteria, visible rows and recalculated subtotals without losing unrelated payloads`, async () => {
    const book = new ExcelJS.Workbook(),
      ws = book.addWorksheet('Data');
    ws.addRows([
      ['Amount', 'Group'],
      [2, 'keep'],
      [4, 'drop'],
      [6, 'keep'],
    ]);
    ws.getCell('D1').value = { formula: 'SUBTOTAL(9,A2:A4)', result: 12 };
    if (kind === 'table')
      ws.addTable({
        name: 'Records',
        ref: 'A1',
        headerRow: true,
        columns: [{ name: 'Amount' }, { name: 'Group' }],
        rows: [
          [2, 'keep'],
          [4, 'drop'],
          [6, 'keep'],
        ],
      });
    const input = new Uint8Array(await book.xlsx.writeBuffer()).buffer;
    const content = await importWorkbook(input);
    if (kind === 'new') {
      delete content.sheets[0].sourcePath;
      delete content.sheets[0].autoFilters;
    }
    const sheet = content.sheets[0],
      filter = sheet.autoFilters?.[0] || {
        ref: 'A1:B4',
        sourcePath: sheet.sourcePath,
        columns: [],
      };
    content.sheets[0] = applySheetFilters(sheet, content.sheets, [
      { ...filter, columns: [{ col: 1, values: ['keep'] }] },
    ]);
    expect(contentSchema.parse(content)).toEqual(content);
    const file = newFile('excel', 'Filtered', content);
    if (kind !== 'new') file.original = { name: 'Filtered.xlsx', data: input };
    const output = await bytes(await exportOffice(file)),
      reimport = await importWorkbook(output);
    expect(reimport.sheets[0].hiddenRows).toEqual([2]);
    expect(reimport.sheets[0].cells.D1.cachedValue).toBe(8);
    expect(reimport.sheets[0].autoFilters?.[0].columns[0]).toMatchObject({
      col: 1,
      values: ['keep'],
    });
    const zip = await JSZip.loadAsync(output),
      original = await JSZip.loadAsync(input);
    if (kind !== 'new')
      for (const path of Object.keys(original.files))
        if (
          !original.files[path].dir &&
          ![
            'xl/worksheets/sheet1.xml',
            'xl/workbook.xml',
            ...(kind === 'table' ? ['xl/tables/table1.xml'] : []),
          ].includes(path)
        )
          expect(await zip.file(path)!.async('uint8array'), path).toEqual(
            await original.file(path)!.async('uint8array'),
          );
    const xml = parseXML(await zip.file('xl/worksheets/sheet1.xml')!.async('string'));
    expect(elements(xml, 'sheetPr')[0].getAttribute('filterMode')).toBe('1');
    const current = file.content as { kind: 'excel'; sheets: Sheet[] };
    current.sheets[0] = applySheetFilters(
      current.sheets[0],
      current.sheets,
      current.sheets[0].autoFilters!.map((f) => ({ ...f, columns: [] })),
    );
    const cleared = await importWorkbook(await bytes(await exportOffice(file)));
    expect(cleared.sheets[0].hiddenRows).toEqual([]);
    expect(cleared.sheets[0].cells.D1.cachedValue).toBe(12);
  });
