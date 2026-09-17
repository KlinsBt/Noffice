import ExcelJS from 'exceljs';
import JSZip from 'jszip';

export async function tableFixture() {
  const book = new ExcelJS.Workbook(),
    ws = book.addWorksheet('Sales data');
  ws.addTable({
    name: 'Sales',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: [{ name: 'Region' }, { name: 'Quantity' }, { name: 'Price' }],
    rows: [
      ['North', 2, 10],
      ['South', 3, 20],
      ['West', 5, 30],
    ],
  });
  ws.getCell('A5').value = 'East';
  ws.getCell('B5').value = 7;
  ws.getCell('C5').value = 40;
  ws.getCell('A6').value = 'Central';
  ws.getCell('B6').value = 11;
  ws.getCell('C6').value = 50;
  ws.getCell('E1').value = { formula: 'SUM(Sales[Quantity])', result: 10 };
  ws.getCell('E2').value = { formula: 'E1*2', result: 20 };
  ws.getCell('E3').value = 'Sales[Quantity]';
  ws.getCell('F1').value = 'Retained note';
  ws.getCell('F1').note = 'Keep this note';
  ws.getCell('F3').value = 'Item';
  ws.getCell('G3').value = 'Amount';
  ws.getCell('F4').value = 'Paper';
  ws.getCell('G4').value = 12;
  ws.getCell('F5').value = 'Ink';
  ws.getCell('G5').value = 25;
  ws.getCell('F6').value = 'Pens';
  ws.getCell('G6').value = 3;
  ws.getCell('B3').numFmt = '0.00';
  for (let c = 1; c <= 7; c++) ws.getColumn(c).width = 19;
  book.addWorksheet('Untouched').getCell('A1').value = 'Keep this sheet';
  const zip = await JSZip.loadAsync(await book.xlsx.writeBuffer());
  zip.file(
    'xl/workbook.xml',
    (await zip.file('xl/workbook.xml')!.async('string')).replace(/ defaultThemeVersion="\d+"/, ''),
  );
  // ExcelJS emits legacyDrawing after tableParts; SpreadsheetML requires the reverse order.
  const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheetXml.replace(/(<tableParts[^>]*>[\s\S]*?<\/tableParts>)(<legacyDrawing[^>]*\/>)/, '$2$1'),
  );
  zip.file('custom/preservation.xml', '<retained>Table editing fixture</retained>');
  return zip.generateAsync({ type: 'arraybuffer' });
}
