import { expect, it } from 'vitest';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import { importWorkbook } from './xlsx-import';
import { calculator } from './formulas';
import native from '../tests/fixtures/native-excel-scoped-names.json';
import { importFile, exportOffice } from './formats';
import { hydrateTableMetadata } from './workbook-tables';
import { contentSchema, type WorkbookContent } from './model';
import JSZip from 'jszip';
const input = () => {
  const b = fs.readFileSync('tests/fixtures/excel-scoped-names.xlsx');
  expect(createHash('sha256').update(b).digest('hex')).toBe(native.sourceSha256);
  return Uint8Array.from(b).buffer;
};
it('matches all native scoped names, workbook aliases and internal workbook qualifiers', async () => {
  const book = await importWorkbook(input()),
    calc = calculator(book.sheets);
  for (const cell of native.cells) {
    const sheet = book.sheets.find((s) => s.name === cell.sheet)!;
    expect(calc.result(sheet, cell.address), `${cell.sheet}!${cell.address}`).toEqual({
      kind: 'value',
      value: cell.value,
    });
  }
});
it('retains typed scope identities and migrates missing metadata without changing cell edits or revision', async () => {
  const data = input();
  const file = await importFile(
    Object.assign(new File([data], 'Names.xlsx'), { arrayBuffer: async () => data }),
  );
  const content = file.content as WorkbookContent;
  expect(
    content.sheets[0].nameDefinitions
      ?.filter((d) => d.name === 'Rate')
      .map((d) => d.scope)
      .sort(),
  ).toEqual(['workbook', 'worksheet']);
  expect(contentSchema.parse(JSON.parse(JSON.stringify(content)))).toEqual(content);
  for (const sheet of content.sheets) delete sheet.nameDefinitions;
  content.sheets[0].cells.A1.value = '5';
  const restored = await hydrateTableMetadata(file);
  expect(restored.revision).toBe(file.revision);
  expect(restored.original!.data).toBe(data);
  const sheets = (restored.content as WorkbookContent).sheets;
  expect(calculator(sheets)(sheets[0], 'C1')).toBe(15);
  content.sheets[0].definedNames!.rate = '99';
  await expect(hydrateTableMetadata(file)).rejects.toThrow('Saved name definitions');
});
it('preserves native scopes and writes recalculated alias caches into the actual XLSX', async () => {
  const data = input();
  const file = await importFile(
    Object.assign(new File([data], 'Names.xlsx'), { arrayBuffer: async () => data }),
  );
  (file.content as WorkbookContent).sheets[0].cells.A1.value = '5';
  const blob = await exportOffice(file);
  const output = await new Promise<ArrayBuffer>((resolve) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as ArrayBuffer);
    r.readAsArrayBuffer(blob);
  });
  const zip = await JSZip.loadAsync(output),
    before = await JSZip.loadAsync(data);
  const names = (xml: string) =>
    new DOMParser()
      .parseFromString(xml, 'application/xml')
      .getElementsByTagNameNS('*', 'definedNames')[0];
  expect(
    names(await zip.file('xl/workbook.xml')!.async('string')).isEqualNode(
      names(await before.file('xl/workbook.xml')!.async('string')),
    ),
  ).toBe(true);
  for (let i = 1; i <= 3; i++) {
    const doc = new DOMParser().parseFromString(
      await zip.file(`xl/worksheets/sheet${i}.xml`)!.async('string'),
      'application/xml',
    );
    const cell = Array.from(doc.getElementsByTagNameNS('*', 'c')).find(
      (e) => e.getAttribute('r') === 'C1',
    )!;
    expect(cell.getElementsByTagNameNS('*', 'v')[0].textContent).toBe('15');
  }
  (file.content as WorkbookContent).sheets[0].nameDefinitions![0].formula = '99';
  await expect(exportOffice(file)).rejects.toThrow('nameDefinitions');
});
it('distinguishes local-to-global same-name aliases, cycles, missing and external names', async () => {
  const book = await importWorkbook(input()),
    sheet = book.sheets[0];
  sheet.nameDefinitions = sheet.nameDefinitions!.map((d) =>
    d.scope === 'worksheet' ? { ...d, formula: '[0]!Rate' } : d,
  );
  let calc = calculator(book.sheets);
  expect(calc(sheet, 'B1')).toBe(2);
  expect(calc.expression(sheet, '[1]!Rate').kind).toBe('unsupported');
  expect(calc.expression(sheet, '[0]!Unknown').value).toBe('#NAME?');
  sheet.nameDefinitions = sheet.nameDefinitions.map((d) =>
    d.scope === 'workbook' && d.name === 'Rate' ? { ...d, formula: 'AliasRate' } : d,
  );
  calc = calculator(book.sheets);
  expect(calc(sheet, 'C1')).toBe('#CYCLE!');
});
it('invalidates workbook alias consumers while retaining unrelated local-name values across revisions', async () => {
  const book = await importWorkbook(input());
  let calc = calculator(book.sheets);
  for (const s of book.sheets) for (const ref of ['B1', 'C1', 'F1']) calc(s, ref);
  const next = book.sheets.map((s, i) =>
    i ? s : { ...s, cells: { ...s.cells, A1: { ...s.cells.A1, value: '5' } } },
  );
  calc = calc.update(next);
  expect(calc(next[0], 'B1')).toBe(11);
  expect(calc(next[1], 'B1')).toBe(7);
  for (const s of next) expect(calc(s, 'C1')).toBe(15);
  expect(calc(next[0], 'F1')).toBe(26);
  expect(calc(next[1], 'F1')).toBe(22);
  expect(calc(next[2], 'F1')).toBe(20);
  calc = calc.update(book.sheets);
  expect(calc(book.sheets[0], 'C1')).toBe(6);
});
