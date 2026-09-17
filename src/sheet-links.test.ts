import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { sheetLinksFixture } from '../tests/fixtures/xlsx-links';
import { importFile, exportOffice } from './formats';
import { importWorkbook, parseXML, elements } from './xlsx-import';
import {
  editSheetLink,
  internalSheetLink,
  removeSheetLinks,
  sheetLinkAddress,
} from './sheet-links';
import type { WorkbookContent } from './model';
async function bytes(blob: Blob) {
  return new Promise<ArrayBuffer>((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as ArrayBuffer);
    r.onerror = reject;
    r.readAsArrayBuffer(blob);
  });
}
async function fixture() {
  const input = await sheetLinksFixture();
  const file = await importFile(
    Object.assign(new File([input], 'Links.xlsx'), { arrayBuffer: async () => input }),
  );
  return { input, file, content: file.content as WorkbookContent };
}
it.each([
  'javascript:alert(1)',
  'data:text/html,test',
  'file:///C:/test',
  '//example.com',
  'https://',
  'https://example.com/ space',
  'mailto:',
  'https://example.com/\n',
])('rejects executable or malformed new addresses %s', (href) => {
  // Leading/trailing user whitespace is trimmed; embedded controls remain invalid.
  if (href.endsWith('\n')) expect(sheetLinkAddress(href)).toBe('https://example.com/');
  else expect(() => sheetLinkAddress(href)).toThrow();
});
it('reads internal locations, ScreenTips and range links from original XML', async () => {
  const { content } = await fixture();
  const sheet = content.sheets[0];
  expect(sheet.cells.B1.hyperlink).toBe("#'O''Brien Data'!$B$3");
  expect(sheet.cells.A1.hyperlinkTooltip).toBe('Original tip');
  expect(sheet.cells.C2.hyperlink).toBe('https://example.com/original');
  expect(internalSheetLink(sheet.cells.B1.hyperlink!, sheet, content.sheets)).toMatchObject({
    sheetId: content.sheets[1].id,
    ref: 'B3',
    end: 'B3',
  });
  expect(() => internalSheetLink('#missing!A1', sheet, content.sheets)).toThrow();
  expect(() => internalSheetLink('#A999999', sheet, content.sheets)).toThrow();
  expect(
    internalSheetLink(
      '#Named',
      { ...sheet, definedNames: { named: "'O''Brien Data'!$B$3:$C$4" } },
      content.sheets,
    ),
  ).toMatchObject({ ref: 'B3', end: 'C4' });
});
it('keeps number/formula values and formatting while changing or removing destinations', async () => {
  const { content } = await fixture();
  const sheet = content.sheets[0];
  const number = editSheetLink(sheet, 'A2', 'mailto:hello@example.com', '42');
  expect(number.cells.A2).toMatchObject({ value: '42', dataType: 'number' });
  const formula = editSheetLink(number, 'A3', '#A2', 'changed label');
  expect(formula.cells.A3).toMatchObject({ value: '=A2*2', cachedValue: 84 });
  const removed = editSheetLink(sheet, 'A1', null, '');
  expect(removed.cells.A1).toMatchObject({
    value: 'Original label',
    bold: true,
    fontSize: 14,
    note: 'Retained note',
  });
  expect(removed.cells.A1.hyperlink).toBeUndefined();
  expect(() => editSheetLink(content.sheets[2], 'A1', '#Links!A1', 'Locked')).toThrow(/protected/);
  expect(removeSheetLinks(content, sheet, ['A2'])).toBe(content);
});
it('patches shared and ranged links without losing other cells or ZIP parts', async () => {
  const { input, file, content } = await fixture();
  let sheet = content.sheets[0];
  sheet = editSheetLink(
    sheet,
    'A1',
    'https://example.com/edited?x=1&y=2',
    'Edited label',
    'Edited tip',
  );
  sheet = editSheetLink(sheet, 'A2', 'mailto:hello@example.com', '42');
  sheet = editSheetLink(sheet, 'A3', '#A2', '=A2*2');
  sheet = editSheetLink(sheet, 'C1', null, '');
  sheet = editSheetLink(sheet, 'E1', '#A2', '=literal');
  content.sheets[0] = sheet;
  const output = await bytes(await exportOffice(file)),
    read = await importWorkbook(output);
  expect(read.sheets[0].cells.A1).toMatchObject({
    hyperlink: 'https://example.com/edited?x=1&y=2',
    hyperlinkTooltip: 'Edited tip',
    value: 'Edited label',
    fontSize: 14,
  });
  expect(read.sheets[0].cells.A2).toMatchObject({ value: '42', dataType: 'number' });
  expect(read.sheets[0].cells.A3).toMatchObject({ value: '=A2*2', hyperlink: '#A2' });
  expect(read.sheets[0].cells.E1).toMatchObject({ value: '=literal', dataType: 'text' });
  expect(read.sheets[0].cells.C1.hyperlink).toBeUndefined();
  for (const ref of ['A4', 'C2'])
    expect(read.sheets[0].cells[ref].hyperlink).toBe('https://example.com/original');
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (
      !before.files[path].dir &&
      ![
        'xl/styles.xml',
        'xl/worksheets/sheet1.xml',
        'xl/worksheets/_rels/sheet1.xml.rels',
        'xl/workbook.xml',
      ].includes(path)
    )
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const xml = parseXML(await after.file('xl/worksheets/sheet1.xml')!.async('string'));
  expect(elements(xml, 'hyperlink').find((e) => e.getAttribute('ref') === 'C2:C2')).toBeTruthy();
});
it('exports formula and numeric hyperlinks in new workbooks without converting their values to strings', async () => {
  const { file, content } = await fixture();
  delete file.original;
  content.sheets = [content.sheets[0], content.sheets[1]];
  delete content.sheets[0].cells.D1.hyperlink;
  content.sheets[0] = editSheetLink(content.sheets[0], 'A2', 'mailto:hello@example.com', '42');
  content.sheets[0] = editSheetLink(content.sheets[0], 'A3', '#A2', '=A2*2');
  const read = await importWorkbook(await bytes(await exportOffice(file)));
  expect(read.sheets[0].cells.A2).toMatchObject({
    value: '42',
    dataType: 'number',
    hyperlink: 'mailto:hello@example.com',
  });
  expect(read.sheets[0].cells.A3).toMatchObject({
    value: '=A2*2',
    cachedValue: 84,
    hyperlink: '#A2',
  });
  expect(read.sheets[0].cells.B1.hyperlink).toBe("#'O''Brien Data'!$B$3");
});
