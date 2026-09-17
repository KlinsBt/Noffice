import ExcelJS from 'exceljs';
import JSZip from 'jszip';
export async function sheetLinksFixture() {
  const book = new ExcelJS.Workbook(),
    sheet = book.addWorksheet('Links');
  sheet.getCell('A1').value = {
    text: 'Original label',
    hyperlink: 'https://example.com/original',
    tooltip: 'Original tip',
  };
  sheet.getCell('A2').value = 42;
  sheet.getCell('A3').value = { formula: 'A2*2', result: 84 };
  sheet.getCell('A4').value = {
    text: 'Keep shared link',
    hyperlink: 'https://example.com/original',
  };
  sheet.getCell('B1').value = 'Internal destination';
  sheet.getCell('C1').value = 'Range one';
  sheet.getCell('C2').value = 'Range two';
  sheet.getCell('D1').value = 'Unsafe imported link';
  sheet.getCell('E1').value = 'Leading equals';
  sheet.getCell('A1').font = { name: 'Arial', size: 14, bold: true, color: { argb: 'FF663399' } };
  sheet.getCell('A1').note = 'Retained note';
  sheet.pageSetup = { printArea: 'A1:E4', orientation: 'landscape' };
  const target = book.addWorksheet("O'Brien Data");
  target.getCell('B3').value = 'Destination';
  const locked = book.addWorksheet('Protected');
  locked.getCell('A1').value = 'Locked';
  await locked.protect('test', { selectLockedCells: true });
  const zip = await JSZip.loadAsync(await book.xlsx.writeBuffer());
  const path = 'xl/worksheets/sheet1.xml';
  const source = await zip.file(path)!.async('string');
  const id = /<hyperlink[^>]*ref="A1"[^>]*r:id="([^"]+)"/.exec(source)![1];
  zip.file(
    path,
    source.replace(
      '</hyperlinks>',
      `<hyperlink ref="B1" location="'O''Brien Data'!$B$3" tooltip="Jump to data"/><hyperlink ref="C1:C2" r:id="${id}"/><hyperlink ref="D1" r:id="rIdUnsafe"/></hyperlinks>`,
    ),
  );
  const relPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  zip.file(
    relPath,
    (await zip.file(relPath)!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rIdUnsafe" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/></Relationships>',
    ),
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}
