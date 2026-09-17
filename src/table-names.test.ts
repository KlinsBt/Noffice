import { it, expect, beforeAll, afterAll } from 'vitest';
import JSZip from 'jszip';
import { newFile } from './model';
import { createSheetTable, validTableName } from './sheet-tables';
import { calculator, translateFormula } from './formulas';
import { exportOffice } from './formats';
import { importWorkbook } from './xlsx-import';
import { renameXlsxTable } from './xlsx-table-rename';
import { tableRenameFixture } from '../tests/fixtures/xlsx-table-rename';
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
it.each(['Ümsatz', '日本', '\\Sales', 'Sales\\West', 'Таблица', '銷售A1'])(
  'creates, calculates and exports table %s',
  async (name) => {
    const file = newFile('excel');
    if (file.content.kind !== 'excel') throw Error('Workbook');
    const s = file.content.sheets[0];
    s.cells = {
      A1: { value: 'Quantity' },
      A2: { value: '4' },
      C1: { value: `=SUM(${name}[Quantity])` },
      C2: { value: `=SUM(INDIRECT("${name}[Quantity]"))` },
    };
    file.content = createSheetTable(file.content, s.id, 'A1:A2', name);
    expect(calculator(file.content.sheets)(file.content.sheets[0], 'C1')).toBe(4);
    expect(calculator(file.content.sheets)(file.content.sheets[0], 'C2')).toBe(4);
    const parsed = await importWorkbook(await (await exportOffice(file)).arrayBuffer());
    expect(parsed.sheets[0].tables![0].name).toBe(name);
    expect(calculator(parsed.sheets)(parsed.sheets[0], 'C1')).toBe(4);
  },
);
it('copies actual A1 references without changing Unicode, dotted or backslash identifiers', () => {
  expect(translateFormula('=SUM(銷售A1[Quantity])+Ümsatz.A1+\\A1+B2+$C$3', 1, 1)).toBe(
    '=SUM(銷售A1[Quantity])+Ümsatz.A1+\\A1+C3+$C$3',
  );
  expect(translateFormula("='日本'!A1+日本!B2", 1, 1)).toBe("='日本'!B2+日本!C3");
});
it.each(['RC', 'R1C', 'RC1', 'R1C1', '日本 table', '1日本', '日本!'])(
  'rejects ambiguous or invalid name %s',
  (name) => {
    expect(() => validTableName(name, [])).toThrow();
  },
);
it('checks Unicode collisions case insensitively', () => {
  const file = newFile('excel');
  if (file.content.kind !== 'excel') throw Error('Workbook');
  const s = file.content.sheets[0];
  s.tables = [{ name: 'Ümsatz', ref: 'A1:B2', headerRows: 1, totalRows: 0, columns: ['A', 'B'] }];
  expect(() => validTableName('üMSATZ', [s])).toThrow(/already used/);
});
it('retains a distinct internal table name during a column-only rename', async () => {
  const zip = await JSZip.loadAsync(await tableRenameFixture());
  const path = 'xl/tables/table1.xml';
  zip.file(
    path,
    (await zip.file(path)!.async('string')).replace('name="Sales"', 'name="InternalSales"'),
  );
  const output = await JSZip.loadAsync(
    await renameXlsxTable(await zip.generateAsync({ type: 'arraybuffer' }), {
      sourcePath: path,
      table: 'Sales',
      name: 'Sales',
      columns: [{ before: 'Quantity', after: 'Units' }],
    }),
  );
  const xml = await output.file(path)!.async('string');
  expect(xml).toContain('name="InternalSales"');
  expect(xml).toContain('displayName="Sales"');
});
