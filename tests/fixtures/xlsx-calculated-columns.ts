import JSZip from 'jszip';
import { tableFixture } from './xlsx-tables';

/** Authored empty Price column, styled cells, a hidden row and dependent totals. */
export async function calculatedColumnsFixture() {
  const zip = await JSZip.loadAsync(await tableFixture());
  let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  for (const r of [2, 3, 4])
    sheet = sheet.replace(new RegExp(`<c r="C${r}"[^>]*>.*?</c>`), `<c r="C${r}"/>`);
  sheet = sheet.replace('<row r="4"', '<row hidden="1" r="4"');
  sheet = sheet.replace(
    /<c r="E1"[^>]*>.*?<\/c>/,
    '<c r="E1"><f>SUM(Sales[Price])</f><v>0</v></c>',
  );
  zip.file('xl/worksheets/sheet1.xml', sheet);
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
