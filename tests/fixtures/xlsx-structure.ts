import ExcelJS from 'exceljs';
import JSZip from 'jszip';

export async function structureFixture(): Promise<ArrayBuffer> {
  const book = new ExcelJS.Workbook();
  const data = book.addWorksheet('Data');
  const other = book.addWorksheet("O'Brien");
  const untouched = book.addWorksheet('Untouched');
  data.getCell('A1').value = 'Amount';
  for (let row = 2; row <= 6; row++) data.getCell(`A${row}`).value = (row - 1) * 10;
  data.getCell('B2').value = { formula: 'A2+$A$4', result: 40 };
  data.getCell('B3').value = { formula: 'SUM(A2:A5)', result: 100 };
  data.getCell('C1').value = { formula: 'SUM(A2:A6)', result: 150 };
  data.getCell('D1').value = { formula: 'A3', result: 20 };
  data.getCell('B5').value = { formula: 'A5*2', result: 80 };
  data.getCell('B6').value = { sharedFormula: 'B5', result: 100 };
  data.getCell('A2').numFmt = '0.00';
  data.getCell('A2').font = { bold: true, color: { argb: 'FF123456' } };
  data.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };
  data.getRow(2).height = 27;
  data.getRow(5).hidden = true;
  data.getColumn(1).width = 18;
  data.getColumn(2).width = 24;
  data.getColumn(5).hidden = true;
  data.mergeCells('E3:F3');
  data.getCell('E3').value = 'Merged heading';
  data.mergeCells('E5:F6');
  data.getCell('E5').value = 'Retained merge';
  for (let row = 2; row <= 6; row++)
    data.getCell(`A${row}`).dataValidation = {
      type: 'whole',
      operator: 'between',
      formulae: [0, 100],
      allowBlank: true,
      showErrorMessage: true,
    };
  data.autoFilter = 'A1:B6';
  data.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
  data.pageSetup = { orientation: 'landscape', printArea: 'A1:F6', printTitlesRow: '1:1' };
  book.definedNames.add('Data!$A$2:$A$6', 'Amounts');
  other.getCell('A1').value = { formula: 'Data!A3', result: 20 };
  other.getCell('A2').value = { formula: 'Data!$B$2', result: 40 };
  other.getCell('A3').value = { formula: 'SUM(Amounts)', result: 150 };
  other.getCell('A4').value = 'A3 is text';
  other.getCell('A5').value = {
    text: 'Jump',
    hyperlink: "#'Data'!$A$5",
    tooltip: 'Jump to amount',
  };
  untouched.getCell('A1').value = 'Keep my package';
  const zip = await JSZip.loadAsync(await book.xlsx.writeBuffer());
  const linksPath = 'xl/worksheets/sheet2.xml';
  zip.file(
    linksPath,
    (await zip.file(linksPath)!.async('string')).replace(
      /(<hyperlink\s[^>]*?)r:id="[^"]+"/,
      '$1location="&apos;Data&apos;!$A$5"',
    ),
  );
  // Retain an unknown but inert package payload to detect fallback rebuilding.
  zip.file(
    'custom/preservation.xml',
    '<preserved>This payload must remain byte-identical.</preserved>',
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}
