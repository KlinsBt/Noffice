import { expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { calculator } from './formulas';
import { importWorkbook, parseXML, elements } from './xlsx-import';
import { exportRetainedWorkbook } from './xlsx-preserve';
import type { OfficeFile } from './model';

it('preserves filter visibility, updates subtotal caches after hiding rows and retains unrelated parts', async () => {
  const book = new ExcelJS.Workbook(),
    ws = book.addWorksheet('Data');
  ws.getCell('A1').value = 'Amount';
  ws.getCell('A2').value = 2;
  ws.getCell('A3').value = 4;
  ws.getCell('A4').value = 6;
  ws.getCell('B1').value = { formula: 'SUBTOTAL(9,A2:A4)', result: 12 };
  ws.getCell('C1').value = { formula: 'SUBTOTAL(109,A2:A4)', result: 12 };
  ws.getCell('D1').value = { formula: 'C1*2', result: 24 };
  const input = new Uint8Array(await book.xlsx.writeBuffer()).buffer;
  const content = await importWorkbook(input);
  content.sheets[0].hiddenRows = [2];
  const file = { content, original: { data: input } } as OfficeFile;
  const output = await exportRetainedWorkbook(file);
  const bytes = await new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = reject;
    reader.readAsArrayBuffer(output);
  });
  const after = await JSZip.loadAsync(bytes),
    before = await JSZip.loadAsync(input);
  const xml = parseXML(await after.file('xl/worksheets/sheet1.xml')!.async('string'));
  const cache = (ref: string) =>
    elements(
      elements(xml, 'c').find((c) => c.getAttribute('r') === ref)!,
      'v',
    )[0].textContent;
  expect(cache('B1')).toBe('12');
  expect(cache('C1')).toBe('8');
  expect(cache('D1')).toBe('16');
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && !['xl/worksheets/sheet1.xml', 'xl/workbook.xml'].includes(path))
      expect(await after.file(path)!.async('uint8array')).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const source = await before.file('xl/worksheets/sheet1.xml')!.async('string');
  before.file(
    'xl/worksheets/sheet1.xml',
    source
      .replace('<sheetFormatPr', '<sheetPr filterMode="1"/><sheetFormatPr')
      .replace(
        '</worksheet>',
        '<autoFilter ref="A1:A4"><filterColumn colId="0"><filters><filter val="2"/><filter val="6"/></filters></filterColumn></autoFilter></worksheet>',
      )
      .replace('<row r="3"', '<row hidden="1" r="3"'),
  );
  const filtered = await importWorkbook(await before.generateAsync({ type: 'arraybuffer' }));
  expect(filtered.sheets[0].filterMode).toBe(true);
  expect(filtered.sheets[0].autoFilters?.[0].columns[0].values).toEqual(['2', '6']);
  expect(calculator(filtered.sheets)(filtered.sheets[0], 'B1')).toBe(8);
  expect(calculator(filtered.sheets)(filtered.sheets[0], 'C1')).toBe(8);
});
