import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { importWorkbook, elements, parseXML } from './xlsx-import';
import { exportRetainedWorkbook } from './xlsx-preserve';
import { newFile, type OfficeFile } from './model';
import { calculator, displayValue } from './formulas';
import { sheetLayout } from './sheet-layout';

async function bytes(blob: Blob): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = reject;
    reader.readAsArrayBuffer(blob);
  });
}
async function fixture(): Promise<OfficeFile> {
  const book = new ExcelJS.Workbook(),
    sheet = book.addWorksheet('Price form');
  sheet.columns = [{ width: 31 }, { width: 15 }, { width: 24 }, { width: 10 }];
  sheet.getRow(1).height = 48;
  sheet.getRow(6).hidden = true;
  sheet.getCell('A6').value = 'Hidden content';
  sheet.mergeCells('A1:D1');
  sheet.getCell('A1').value = 'Price form';
  sheet.getCell('A1').font = { name: 'Arial', size: 14, bold: true, color: { theme: 0 } };
  sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { theme: 4 } };
  sheet.getCell('A1').alignment = { wrapText: true, vertical: 'middle' };
  sheet.getCell('A1').border = { bottom: { style: 'medium', color: { argb: 'FF123456' } } };
  sheet.getCell('A3').value = '001234';
  sheet.getCell('B3').value = 2;
  sheet.getCell('C3').value = 75;
  sheet.getCell('B3').protection = { locked: false };
  sheet.getCell('C3').numFmt = '#,##0.00 "EUR"';
  sheet.getCell('D3').value = { formula: 'B3*C3', result: 150 };
  sheet.getCell('A4').value = new Date('2025-12-03T00:00:00Z');
  sheet.getCell('A4').numFmt = 'dd.mm.yyyy';
  sheet.getCell('B4').value = { text: 'Details', hyperlink: 'https://example.com/details' };
  sheet.getCell('A5').value = {
    richText: [{ text: 'CO' }, { text: '2', font: { vertAlign: 'subscript' } }],
  };
  sheet.getCell('C5').border = { bottom: { style: 'thin' } };
  sheet.getCell('B3').dataValidation = { type: 'whole', operator: 'greaterThan', formulae: [0] };
  sheet.getCell('B3').note = 'Enter quantity';
  sheet.pageSetup = {
    paperSize: 9,
    orientation: 'landscape',
    printArea: 'A1:D10',
    printTitlesRow: '1:2',
  };
  sheet.headerFooter = { oddFooter: '&P / &N' };
  sheet.views = [{ state: 'frozen', ySplit: 2, showGridLines: false }];
  book.definedNames.add("'Price form'!$C$3", 'Rate');
  const hidden = book.addWorksheet('Lookup', { state: 'hidden' });
  hidden.getCell('A1').value = 'Internal';
  const picture = book.addImage({
    base64:
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==',
    extension: 'png',
  });
  sheet.addImage(picture, 'D5:D6');
  await sheet.protect('test-only', { spinCount: 1 });
  const zip = await JSZip.loadAsync(await book.xlsx.writeBuffer());
  zip.file(
    'xl/calcChain.xml',
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="D3" i="1"/></calcChain>',
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    (await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rIdCalcChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
    ),
  );
  zip.file(
    '[Content_Types].xml',
    (await zip.file('[Content_Types].xml')!.async('string')).replace(
      '</Types>',
      '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
    ),
  );
  zip.file('customXml/keep.xml', '<retained>This part must survive.</retained>');
  const source = await zip.generateAsync({ type: 'arraybuffer' });
  const file = newFile('excel', 'Fixture', await importWorkbook(source));
  file.original = { name: 'fixture.xlsx', data: source };
  return file;
}
describe('preservation of imported XLSX packages', () => {
  it('writes new cells and styles using the original Strict SpreadsheetML namespace', async () => {
    const file = await fixture();
    const zip = await JSZip.loadAsync(file.original!.data);
    const strict = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    for (const part of Object.values(zip.files))
      if (!part.dir && part.name.endsWith('.xml'))
        zip.file(
          part.name,
          (await part.async('string')).replaceAll(
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            strict,
          ),
        );
    file.original!.data = await zip.generateAsync({ type: 'arraybuffer' });
    file.content = await importWorkbook(file.original!.data);
    file.content.sheets[0].cells.B3.value = 'Changed text';
    file.content.sheets[0].cells.C7 = { value: 'New cell', fill: '#123456' };
    const output = await JSZip.loadAsync(await bytes(await exportRetainedWorkbook(file)));
    for (const path of ['xl/worksheets/sheet1.xml', 'xl/styles.xml', 'xl/workbook.xml']) {
      const xml = await output.file(path)!.async('string');
      expect(xml).not.toContain('http://schemas.openxmlformats.org/spreadsheetml/2006/main');
    }
  });
  it('returns byte-identical original data for an unchanged import', async () => {
    const file = await fixture();
    expect(new Uint8Array(await bytes(await exportRetainedWorkbook(file)))).toEqual(
      new Uint8Array(file.original!.data),
    );
  });
  it('retains layout, cell types, styles and hidden state in the editor model', async () => {
    const file = await fixture();
    if (file.content.kind !== 'excel') throw new Error();
    const [s, h] = file.content.sheets;
    expect(s.merges).toEqual(['A1:D1']);
    expect(s.rowHeights?.[0]).toBe(64);
    expect(s.hiddenRows).toContain(5);
    expect(s.columnWidths?.[0]).toBe(222);
    expect(h.state).toBe('hidden');
    expect(s.cells.A1.color).toBe('#FFFFFF');
    expect(s.cells.A1.fill).toMatch(/^#/);
    expect(s.cells.A1.borders?.bottom).toBe('2px solid #123456');
    expect(s.cells.A1.wrap).toBe(true);
    expect(s.frozenRows).toBe(2);
    const calc = calculator(file.content.sheets);
    expect(calc(s, 'A3')).toBe('001234');
    expect(displayValue(calc(s, 'A4'), s.cells.A4)).toBe('03.12.2025');
    expect(displayValue(calc(s, 'C3'), s.cells.C3)).toBe('75.00 EUR');
    const layout = sheetLayout(s);
    expect(layout.rows).not.toContain(5);
    expect(layout.cell(0, 0).colspan).toBe(4);
    expect(layout.cell(0, 2).hidden).toBe(true);
  });
  it('patches an input, recalculates supported formulas and preserves every unrelated part', async () => {
    const file = await fixture();
    if (file.content.kind !== 'excel') throw new Error();
    file.content.sheets[0].cells.B3.value = '4';
    const data = await bytes(await exportRetainedWorkbook(file)),
      before = await JSZip.loadAsync(file.original!.data),
      after = await JSZip.loadAsync(data);
    expect(Object.keys(after.files).sort()).toEqual(Object.keys(before.files).sort());
    for (const name of Object.keys(before.files))
      if (
        !before.files[name].dir &&
        !['xl/worksheets/sheet1.xml', 'xl/workbook.xml'].includes(name)
      )
        expect(await after.file(name)!.async('string'), name).toBe(
          await before.file(name)!.async('string'),
        );
    const book = new ExcelJS.Workbook();
    await book.xlsx.load(data);
    const sheet = book.worksheets[0];
    expect(sheet.getCell('B3').value).toBe(4);
    expect(sheet.getCell('D3').result).toBe(300);
    expect(sheet.getCell('A3').value).toBe('001234');
    expect(sheet.getCell('A4').value).toBeInstanceOf(Date);
    expect(sheet.getCell('A5').value).toHaveProperty('richText');
    expect(sheet.getCell('B4').hyperlink).toBe('https://example.com/details');
    expect(sheet.getImages()).toHaveLength(1);
    expect(sheet.getCell('B3').note).toBe('Enter quantity');
    expect(sheet.pageSetup.orientation).toBe('landscape');
    expect(sheet.pageSetup.printArea).toBe('A1:D10');
    expect(sheet.getCell('B3').dataValidation.formulae).toEqual([0]);
  });
  it('changes one font property without deleting the source theme fill or border', async () => {
    const file = await fixture();
    if (file.content.kind !== 'excel') throw new Error();
    file.content.sheets[0].cells.A1.fontFamily = 'Georgia';
    const data = await bytes(await exportRetainedWorkbook(file)),
      book = new ExcelJS.Workbook();
    await book.xlsx.load(data);
    const cell = book.worksheets[0].getCell('A1');
    expect(cell.font.name).toBe('Georgia');
    expect(cell.font.color).toEqual({ theme: 0 });
    expect(cell.fill).toMatchObject({ fgColor: { theme: 4 } });
    expect(cell.border.bottom?.style).toBe('medium');
    expect(cell.alignment.wrapText).toBe(true);
  });
  it('exports resized geometry without dropping merges, hidden rows or worksheet protection', async () => {
    const file = await fixture();
    if (file.content.kind !== 'excel') throw new Error();
    const sheet = file.content.sheets[0];
    sheet.columnWidths![1] = 215;
    sheet.rowHeights![0] = 80;
    const data = await bytes(await exportRetainedWorkbook(file)),
      book = new ExcelJS.Workbook();
    await book.xlsx.load(data);
    const s = book.worksheets[0];
    expect(s.getColumn(2).width).toBe(30);
    expect(s.getRow(1).height).toBe(60);
    expect(s.getRow(6).hidden).toBe(true);
    expect(s.getCell('B1').isMerged).toBe(true);
    const zip = await JSZip.loadAsync(data);
    expect(
      elements(parseXML(await zip.file(sheet.sourcePath!)!.async('string')), 'sheetProtection'),
    ).toHaveLength(1);
  });
  it('does not replace unsupported formula caches with IFERROR fallback during unrelated edits', async () => {
    const file = await fixture();
    if (file.content.kind !== 'excel') throw new Error();
    const source = await JSZip.loadAsync(file.original!.data),
      path = file.content.sheets[0].sourcePath!;
    const doc = parseXML(await source.file(path)!.async('string'));
    const cell = elements(doc, 'c').find((e) => e.getAttribute('r') === 'D3')!;
    elements(cell, 'f')[0].textContent = 'IFERROR(FUTUREFUNCTION(B3),"wrong fallback")';
    cell.setAttribute('t', 'str');
    elements(cell, 'v')[0].textContent = 'original cached result';
    source.file(path, new XMLSerializer().serializeToString(doc));
    file.original!.data = await source.generateAsync({ type: 'arraybuffer' });
    file.content = await importWorkbook(file.original!.data);
    file.content.sheets[0].cells.B3.value = '9';
    const after = await JSZip.loadAsync(await bytes(await exportRetainedWorkbook(file)));
    const updated = elements(parseXML(await after.file(path)!.async('string')), 'c').find(
      (e) => e.getAttribute('r') === 'D3',
    )!;
    expect(elements(updated, 'v')[0].textContent).toBe('original cached result');
  });
  it('keeps formula-returned error-looking text as text, not an Excel error cell', async () => {
    const file = await fixture();
    if (file.content.kind !== 'excel') throw new Error();
    file.content.sheets[0].cells.D3.value = '=IF(B3>0,"#N/A","ok")';
    const zip = await JSZip.loadAsync(await bytes(await exportRetainedWorkbook(file)));
    const node = elements(
      parseXML(await zip.file(file.content.sheets[0].sourcePath!)!.async('string')),
      'c',
    ).find((e) => e.getAttribute('r') === 'D3')!;
    expect(node.getAttribute('t')).toBe('str');
    expect(elements(node, 'v')[0].textContent).toBe('#N/A');
    expect(zip.file('xl/calcChain.xml')).toBeNull();
    expect(await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).not.toContain(
      '/calcChain',
    );
    expect(await zip.file('[Content_Types].xml')!.async('string')).not.toContain(
      '/xl/calcChain.xml',
    );
  });
});
