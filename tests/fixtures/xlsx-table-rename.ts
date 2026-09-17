import JSZip from 'jszip';
import { tableFixture } from './xlsx-tables';

/** Authored dependency fixture; no third-party workbook contents. */
export async function tableRenameFixture() {
  const zip = await JSZip.loadAsync(await tableFixture());
  let sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  const formula = (ref: string, text: string, value: string | number) =>
    `<c r="${ref}"${typeof value === 'string' ? ' t="str"' : ''}><f>${text}</f><v>${value}</v></c>`;
  for (const [r, value] of [
    [2, 20],
    [3, 30],
    [4, 50],
  ])
    sheet = sheet.replace(
      new RegExp(`<c r="C${r}"[^>]*>.*?</c>`),
      formula(`C${r}`, 'Sales[[#This Row],[Quantity]]*10', value),
    );
  const extras = [
    formula('H1', 'SUM(Sales[#Data])', 110),
    formula('H2', 'SUM(Units)', 10),
    formula('H3', 'Sales[[#Headers],[Quantity]]', 'Quantity'),
    formula('H4', 'SUM(Sales[[Quantity]:[Price]])', 110),
    formula('H5', '&quot;Sales[Quantity]&quot;', 'Sales[Quantity]'),
    formula('H6', 'SUM(INDIRECT(&quot;Sales[Quantity]&quot;))', 10),
  ];
  extras.forEach((cell, i) => {
    sheet = sheet.replace(new RegExp(`(<row r="${i + 1}"[^>]*>.*?)</row>`), '$1' + cell + '</row>');
  });
  sheet = sheet.replace(
    '<pageMargins',
    '<conditionalFormatting sqref="J2:J5"><cfRule type="expression" priority="1"><formula>SUM(Units)&gt;0</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="whole" operator="lessThanOrEqual" sqref="K2:K5"><formula1>SUM(Units)</formula1></dataValidation></dataValidations><pageMargins',
  );
  zip.file('xl/worksheets/sheet1.xml', sheet);
  zip.file(
    'xl/workbook.xml',
    (await zip.file('xl/workbook.xml')!.async('string')).replace(
      '<calcPr',
      '<definedNames><definedName name="Units">Sales[Quantity]</definedName></definedNames><calcPr',
    ),
  );
  const path = Object.keys(zip.files).find((p) => /^xl\/tables\/[^/]+\.xml$/.test(p))!;
  zip.file(
    path,
    (await zip.file(path)!.async('string')).replace(
      /<tableColumn id="3"[^>]*\/>/,
      '<tableColumn id="3" name="Price"><calculatedColumnFormula>Sales[[#This Row],[Quantity]]*10</calculatedColumnFormula></tableColumn>',
    ),
  );
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
