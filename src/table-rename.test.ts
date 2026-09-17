import { it, expect, beforeAll, afterAll } from 'vitest';
import JSZip from 'jszip';
import { renameTableReferences } from './table-references';
import { renameWorkbookTable } from './workbook-table-rename';
import { renameXlsxTable } from './xlsx-table-rename';
import { tableRenameFixture } from '../tests/fixtures/xlsx-table-rename';
import { importFile, exportOffice } from './formats';
import { elements, parseXML, importWorkbook } from './xlsx-import';
import { calculator } from './formulas';
import { newFile } from './model';
import { createSheetTable } from './sheet-tables';

const native = Blob.prototype.arrayBuffer;
beforeAll(() => {
  if (!native)
    Blob.prototype.arrayBuffer = function () {
      return new Promise((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(r.result as ArrayBuffer);
        r.onerror = reject;
        r.readAsArrayBuffer(this);
      });
    };
});
afterAll(() => {
  if (!native) delete (Blob.prototype as Partial<Blob>).arrayBuffer;
});
const edit = {
  table: 'Sales',
  name: 'Orders',
  column: { before: 'Quantity', after: 'Units sold' },
};
it.each([
  [
    '=SUM(Sales[Quantity])+SUM(sales)+Sales2[Quantity]',
    '=SUM(Orders[Units sold])+SUM(sales)+Sales2[Quantity]',
  ],
  ['=Sales[[#Headers],[Quantity]]', '=Orders[[#Headers],[Units sold]]'],
  ['=Sales[[Quantity]:[Price]]', '=Orders[[Units sold]:[Price]]'],
  ['=Sales[@Quantity]', '=Orders[@Units sold]'],
  ["=Sales!A1+'Sales'!B1+Sales(2)", "=Sales!A1+'Sales'!B1+Sales(2)"],
  ['="Sales[Quantity]"&"He said ""Sales"""', '="Sales[Quantity]"&"He said ""Sales"""'],
  ['=INDIRECT("Sales[Quantity]")', '=INDIRECT("Sales[Quantity]")'],
  ['=Other[Quantity]+Sales[[#All],[Quantity]]', '=Other[Quantity]+Orders[[#All],[Units sold]]'],
])('repairs tokens without replacing unrelated text: %s', (input, expected) => {
  expect(renameTableReferences(input, edit)).toBe(expected);
});
it('repairs unqualified references only in the owning table', () => {
  expect(renameTableReferences('=[@Quantity]*10', edit, 'Sales')).toBe('=[@Units sold]*10');
  expect(renameTableReferences('=[@Quantity]*10', edit, 'Other')).toBe('=[@Quantity]*10');
  expect(() => renameTableReferences('=[@Quantity]', edit)).toThrow(/context/);
  expect(() => renameTableReferences('=SUM([1]Sales[Quantity])', edit)).toThrow(/External/);
  expect(() => renameTableReferences("='[other.xlsx]Sheet1'!Sales[Quantity]", edit)).toThrow(
    /External/,
  );
  expect(() => renameTableReferences('=Sales[Quantity', edit)).toThrow(/brackets/);
});
it('uses the native REF error for a missing INDIRECT table while retaining NAME for a direct missing table', () => {
  const file = newFile('excel');
  if (file.content.kind !== 'excel') throw Error('Workbook expected');
  const s = file.content.sheets[0];
  s.cells = {
    A1: { value: '=SUM(INDIRECT("Missing[Quantity]"))' },
    A2: { value: '=SUM(Missing[Quantity])' },
  };
  const calc = calculator(file.content.sheets);
  expect(calc(s, 'A1')).toBe('#REF!');
  expect(calc(s, 'A2')).toBe('#NAME?');
});
it('escapes special headers and reverses chained renames', () => {
  const special = { ...edit, column: { before: 'Quantity', after: "Net, [#'@]" } };
  const result = renameTableReferences('=Sales[Quantity]+Sales[@Quantity]', special);
  expect(result).toBe("=Orders[[Net, '['#'''@']]]+Orders[@[Net, '['#'''@']]]");
  expect(
    renameTableReferences(result, {
      table: 'Orders',
      name: 'Sales',
      column: { before: special.column.after, after: 'Quantity' },
    }),
  ).toBe('=Sales[[Quantity]]+Sales[@[Quantity]]');
});
async function fixture() {
  const data = await tableRenameFixture();
  const file = await importFile({
    name: 'Rename.xlsx',
    size: data.byteLength,
    arrayBuffer: async () => data,
  } as File);
  if (file.content.kind !== 'excel') throw Error('Workbook expected');
  return { file, data, content: file.content };
}
it('repairs retained formulas, names, rules and calculated columns, preserves original and unrelated parts across consecutive renames', async () => {
  const { file, data, content } = await fixture();
  content.sheets[0].cells.B2.value = '4';
  const next = await renameWorkbookTable(file, {
    sheetId: content.sheets[0].id,
    table: 'Sales',
    name: 'Orders',
    column: { index: 1, name: 'Units sold' },
  });
  expect(content.sheets[0].tables![0].name).toBe('Sales');
  expect(next.sheets[0].id).toBe(content.sheets[0].id);
  expect(next.sheets[0].tables![0].columns[1]).toBe('Units sold');
  expect(next.sheets[0].cells.B1.value).toBe('Units sold');
  const calc = calculator(next.sheets),
    s = next.sheets[0];
  expect(calc(s, 'E1')).toBe(12);
  expect(calc(s, 'H2')).toBe(12);
  expect(calc(s, 'H3')).toBe('Units sold');
  expect(calc(s, 'C2')).toBe(40);
  expect(s.cells.H5.value).toBe('="Sales[Quantity]"');
  const exported = await (await exportOffice({ ...file, content: next })).arrayBuffer();
  const zip = await JSZip.loadAsync(exported),
    original = await JSZip.loadAsync(data);
  for (const path of [
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/comments1.xml',
    'xl/worksheets/sheet2.xml',
    'custom/preservation.xml',
  ])
    expect(await zip.file(path)!.async('string')).toBe(await original.file(path)!.async('string'));
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  expect(xml).toContain('<formula1>SUM(Units)</formula1>');
  expect(xml).toContain('<formula>SUM(Units)&gt;0</formula>');
  const doc = parseXML(xml),
    at = (ref: string) => elements(doc, 'c').find((c) => c.getAttribute('r') === ref)!;
  expect(elements(at('H3'), 'v')[0].textContent).toBe('Units sold');
  expect(at('H6').getAttribute('t')).toBe('e');
  expect(elements(at('H6'), 'v')[0].textContent).toBe('#REF!');
  expect(await zip.file(s.tables![0].sourcePath!)!.async('string')).toContain(
    'Orders[[#This Row],[Units sold]]*10',
  );
  const twice = await renameWorkbookTable(
    { ...file, content: next },
    { sheetId: s.id, table: 'Orders', name: 'Revenue', column: { index: 1, name: 'Net, total' } },
  );
  expect(calculator(twice.sheets)(twice.sheets[0], 'E1')).toBe(12);
  expect(
    (await importWorkbook(await (await exportOffice({ ...file, content: twice })).arrayBuffer()))
      .sheets[0].tables![0].name,
  ).toBe('Revenue');
});
it('rejects invalid names, duplicate headers and protected workbooks without changing the source', async () => {
  const { file, content } = await fixture(),
    sheetId = content.sheets[0].id;
  for (const name of ['A1', 'bad name', 'Units'])
    await expect(renameWorkbookTable(file, { sheetId, table: 'Sales', name })).rejects.toThrow();
  for (const name of ['', 'Price', 'a'.repeat(256)])
    await expect(
      renameWorkbookTable(file, {
        sheetId,
        table: 'Sales',
        name: 'Orders',
        column: { index: 1, name },
      }),
    ).rejects.toThrow();
  content.sheets[0].protected = true;
  await expect(
    renameWorkbookTable(file, { sheetId, table: 'Sales', name: 'Orders' }),
  ).rejects.toThrow(/protected/);
  expect(content.sheets[0].tables![0].name).toBe('Sales');
});
it.each([
  'xl/pivotCache/pivotCacheDefinition1.xml',
  'xl/connections.xml',
  'xl/externalLinks/externalLink1.xml',
  'xl/vbaProject.bin',
])('rejects unhandled dependencies: %s', async (path) => {
  const zip = await JSZip.loadAsync(await tableRenameFixture());
  zip.file(path, '<dependency/>');
  await expect(
    renameXlsxTable(await zip.generateAsync({ type: 'arraybuffer' }), {
      table: 'Sales',
      name: 'Orders',
      sourcePath: 'xl/tables/table1.xml',
    }),
  ).rejects.toThrow(/not supported/);
});
it('renames new workbooks without disabling subsequent sheet creation or table exports', async () => {
  const file = newFile('excel');
  if (file.content.kind !== 'excel') throw Error('Workbook expected');
  const s = file.content.sheets[0];
  s.cells = {
    A1: { value: 'Quantity' },
    A2: { value: '3' },
    C1: { value: '=SUM(Sales[Quantity])' },
  };
  file.content = createSheetTable(file.content, s.id, 'A1:A2', 'Sales');
  const next = await renameWorkbookTable(file, {
    sheetId: s.id,
    table: 'Sales',
    name: 'Orders',
    column: { index: 0, name: 'Units sold' },
  });
  expect(next.xlsxStructureBase).toBeUndefined();
  expect(next.sheets[0].sourcePath).toBeUndefined();
  const result = await importWorkbook(
    await (await exportOffice({ ...file, content: next })).arrayBuffer(),
  );
  expect(calculator(result.sheets)(result.sheets[0], 'C1')).toBe(3);
  expect(result.sheets[0].tables![0].name).toBe('Orders');
});

it('repairs chart formulas and direct rule expressions at the package level without rewriting chart caches', async () => {
  const zip = await JSZip.loadAsync(await tableRenameFixture());
  zip.file(
    'xl/charts/chart1.xml',
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:f>Sales[Quantity]</c:f><c:v>Sales[Quantity]</c:v></c:chart></c:chartSpace>',
  );
  zip.file(
    'xl/worksheets/sheet1.xml',
    (await zip.file('xl/worksheets/sheet1.xml')!.async('string'))
      .replaceAll('<formula1>SUM(Units)', '<formula1>SUM(Sales[Quantity])')
      .replaceAll('<formula>SUM(Units)', '<formula>SUM(Sales[Quantity])'),
  );
  const output = await JSZip.loadAsync(
    await renameXlsxTable(await zip.generateAsync({ type: 'arraybuffer' }), {
      ...edit,
      sourcePath: 'xl/tables/table1.xml',
      columnIndex: 1,
    }),
  );
  expect(await output.file('xl/charts/chart1.xml')!.async('string')).toContain(
    '<c:f>Orders[Units sold]</c:f><c:v>Sales[Quantity]</c:v>',
  );
  const xml = await output.file('xl/worksheets/sheet1.xml')!.async('string');
  expect(xml).toContain('<formula1>SUM(Orders[Units sold])</formula1>');
  expect(xml).toContain('<formula>SUM(Orders[Units sold])&gt;0</formula>');
});
it.each(['extension', 'scoped', 'ambiguous'])(
  'rejects an unrepairable %s reference before returning output',
  async (kind) => {
    const zip = await JSZip.loadAsync(await tableRenameFixture());
    let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    if (kind === 'extension')
      xml = xml.replace('</worksheet>', '<extLst><ext uri="custom"/></extLst></worksheet>');
    if (kind === 'scoped')
      xml = xml.replace('<f>SUM(Sales[Quantity])</f>', '<f>LET(x,Sales[Quantity],SUM(x))</f>');
    if (kind === 'ambiguous')
      xml = xml.replace('<f>SUM(Sales[Quantity])</f>', '<f>SUM([Quantity])</f>');
    zip.file('xl/worksheets/sheet1.xml', xml);
    await expect(
      renameXlsxTable(await zip.generateAsync({ type: 'arraybuffer' }), {
        ...edit,
        sourcePath: 'xl/tables/table1.xml',
        columnIndex: 1,
      }),
    ).rejects.toThrow(/references|context/);
  },
);
