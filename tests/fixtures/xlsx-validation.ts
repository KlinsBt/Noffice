import ExcelJS from 'exceljs';
import JSZip from 'jszip';
export async function validationFixture() {
  const book = new ExcelJS.Workbook(),
    sheet = book.addWorksheet('Entry');
  sheet.getCell('A1').value = 'Validated input';
  sheet.getCell('A2').value = 5;
  sheet.getCell('B2').value = 2;
  sheet.getCell('B3').value = 3;
  sheet.getCell('E1').value = { formula: 'A2*2', result: 10 };
  sheet.getCell('A2').note = 'Preserve this note';
  sheet.pageSetup = { printArea: 'A1:E5', orientation: 'landscape' };
  book.addWorksheet('Other').getCell('A1').value = 'Untouched';
  const locked = book.addWorksheet('Protected');
  locked.getCell('A1').value = 'Locked';
  await locked.protect('test', {});
  const zip = await JSZip.loadAsync(await book.xlsx.writeBuffer()),
    path = 'xl/worksheets/sheet1.xml';
  zip.file(
    path,
    (await zip.file(path)!.async('string')).replace(
      '</sheetData>',
      '</sheetData><dataValidations count="2"><dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" showInputMessage="1" errorStyle="stop" prompt="Enter a whole number" error="Use the allowed range" imeMode="noControl" sqref="A2:A3"><formula1>B2</formula1><formula2>B2+10</formula2></dataValidation><dataValidation type="list" allowBlank="1" showErrorMessage="1" showInputMessage="1" errorStyle="stop" prompt="Choose a color" sqref="D2:D1048576"><formula1>"Red,Green"</formula1></dataValidation></dataValidations>',
    ),
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}
