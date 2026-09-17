import { expect, it } from 'vitest';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import native from '../tests/fixtures/native-excel-relative-names.json';
import { calculator } from './formulas';
import { importWorkbook } from './xlsx-import';
import { importFile, exportOffice } from './formats';
import { hydrateTableMetadata } from './workbook-tables';
import { contentSchema, type WorkbookContent } from './model';

const input = () => {
  const bytes = fs.readFileSync('tests/fixtures/excel-relative-names.xlsx');
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(native.sourceSha256);
  return Uint8Array.from(bytes).buffer;
};
function stage(book: WorkbookContent, name: string) {
  const data = book.sheets.find((s) => s.name === 'Data')!;
  if (name === 'edit') data.cells.A3 = { value: '500' };
  if (name === 'delete' || name === 'caller') delete data.cells.A3;
  if (name === 'caller') {
    const calls = book.sheets.find((s) => s.name === 'Calls')!;
    calls.cells.D2 = { value: '=1/RowNext' };
    calls.cells.D3 = { value: '=RowNext' };
  }
  return book;
}
const expected = (c: { value: unknown; valueType: string }) =>
  c.valueType === 'System.Int32' ? '#DIV/0!' : c.value === null ? '' : c.value;
it('binds saved names from A1 at every caller and matches 196 native source/edit/delete/caller snapshots', async () => {
  const original = await importWorkbook(input());
  let calc = calculator(original.sheets);
  for (const oracle of native.stages) {
    const book = stage(structuredClone(original), oracle.stage);
    calc = calc.update(book.sheets);
    for (const cell of oracle.cells) {
      const sheet = book.sheets.find((s) => s.name === cell.sheet)!;
      expect(
        calc.result(sheet, cell.address),
        `${oracle.stage}: ${cell.sheet}!${cell.address}`,
      ).toEqual({
        kind: cell.valueType === 'System.Int32' ? 'error' : 'value',
        value: expected(cell),
      });
    }
  }
  calc = calc.update(original.sheets);
  expect(calc(original.sheets[1], 'D2')).toBe(301);
});
it('keeps unknown origins, expressions without caller context, external sources and name cycles explicit', async () => {
  const book = await importWorkbook(input());
  const calls = book.sheets[1];
  calls.cells.D3 = { value: '=IFERROR(Area,0)' };
  // The native intersection fixture's RelativeArea case catches the same missing intersection.
  expect(calculator(book.sheets).result(calls, 'D3')).toEqual({ kind: 'value', value: 0 });
  expect(calculator(book.sheets).expression(calls, 'RowNext').value).toBe('#UNSUPPORTED!');
  delete calls.nameDefinitions!.find((d) => d.name === 'RowNext')!.referenceOrigin;
  calls.cells.D2 = { value: '=IFERROR(RowNext,0)' };
  expect(calculator(book.sheets).result(calls, 'D2').kind).toBe('unsupported');
  calls.nameDefinitions!.push({
    name: 'Loop',
    formula: 'Loop',
    scope: 'workbook',
    referenceOrigin: 'ooxml-a1',
  });
  calls.cells.D2 = { value: '=Loop' };
  expect(calculator(book.sheets)(calls, 'D2')).toBe('#CYCLE!');
  calls.cells.D2 = { value: '=IFERROR([1]!RowNext,0)' };
  expect(calculator(book.sheets).result(calls, 'D2').kind).toBe('unsupported');
});
it('recovers origin metadata in legacy saved names without rewriting cells, revision or original bytes', async () => {
  const data = input();
  const file = await importFile(
    Object.assign(new File([data], 'Relative.xlsx'), { arrayBuffer: async () => data }),
  );
  const content = file.content as WorkbookContent;
  expect(contentSchema.parse(JSON.parse(JSON.stringify(content)))).toEqual(content);
  for (const s of content.sheets) for (const d of s.nameDefinitions!) delete d.referenceOrigin;
  content.sheets[0].cells.A3.value = '500';
  const restored = await hydrateTableMetadata(file);
  expect(restored.revision).toBe(file.revision);
  expect(restored.original!.data).toBe(data);
  const sheets = (restored.content as WorkbookContent).sheets;
  expect(calculator(sheets)(sheets[1], 'D2')).toBe(500);
  content.sheets[1].nameDefinitions![0].formula = '99';
  await expect(hydrateTableMetadata(file)).rejects.toThrow('Saved name definitions');
});
it('preserves names and unrelated package payloads while exporting native typed caches', async () => {
  const data = input();
  const file = await importFile(
    Object.assign(new File([data], 'Relative.xlsx'), { arrayBuffer: async () => data }),
  );
  const before = await JSZip.loadAsync(data);
  stage(file.content as WorkbookContent, 'caller');
  const output = await exportOffice(file);
  const bytes = await new Promise<ArrayBuffer>((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.readAsArrayBuffer(output);
  });
  const after = await JSZip.loadAsync(bytes);
  const allowed = new Set([
    'xl/workbook.xml',
    'xl/worksheets/sheet1.xml',
    'xl/worksheets/sheet2.xml',
    'xl/worksheets/sheet3.xml',
    'xl/calcChain.xml',
    '[Content_Types].xml',
    'xl/_rels/workbook.xml.rels',
  ]);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && !allowed.has(path))
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const names = (text: string) =>
    new DOMParser()
      .parseFromString(text, 'application/xml')
      .getElementsByTagNameNS('*', 'definedNames')[0].outerHTML;
  expect(names(await after.file('xl/workbook.xml')!.async('string'))).toBe(
    names(await before.file('xl/workbook.xml')!.async('string')),
  );
  for (const cell of native.stages.find((s) => s.stage === 'caller')!.cells) {
    if (!cell.formula.startsWith('=')) continue;
    const path = cell.sheet === 'Calls' ? 'xl/worksheets/sheet2.xml' : 'xl/worksheets/sheet3.xml';
    const xml = new DOMParser().parseFromString(
      await after.file(path)!.async('string'),
      'application/xml',
    );
    const node = Array.from(xml.getElementsByTagNameNS('*', 'c')).find(
      (c) => c.getAttribute('r') === cell.address,
    )!;
    expect(
      node.getElementsByTagNameNS('*', 'v')[0].textContent,
      `${cell.sheet}!${cell.address}`,
    ).toBe(String(expected(cell)));
    expect(node.getAttribute('t') || 'n').toBe(cell.valueType === 'System.Int32' ? 'e' : 'n');
  }
});
