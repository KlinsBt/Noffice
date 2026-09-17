import { expect, it } from 'vitest';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import native from '../tests/fixtures/native-excel-sumproduct.json';
import { calculator } from './formulas';
import { importWorkbook } from './xlsx-import';
import { importFile, exportOffice } from './formats';
import type { WorkbookContent } from './model';
const input = () => {
  const bytes = fs.readFileSync('tests/fixtures/excel-sumproduct.xlsx');
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(native.sourceSha256);
  return Uint8Array.from(bytes).buffer;
};
it('bounds array work, rejects malformed constants and recovers without changing scalar callers', async () => {
  const book = await importWorkbook(input());
  const calls = book.sheets.find((s) => s.name === 'Calls')!;
  const calc = calculator(book.sheets);
  // Broader function lifting must not silently scalarize new array constants.
  for (const formula of ['SUM({1;2}*2)', 'SUMPRODUCT(SQRT({1;4}))'])
    expect(calc.expression(calls, formula)).toEqual({
      kind: 'unsupported',
      value: '#UNSUPPORTED!',
    });
  for (const formula of ['SUMPRODUCT()', `SUMPRODUCT(${Array(256).fill('1').join(',')})`])
    expect(calc.expression(calls, formula)).toEqual({ kind: 'unsupported', value: '#VALUE!' });
  for (const formula of ['SUMPRODUCT({1,2;3})', 'SUMPRODUCT({A1,2})', 'SUMPRODUCT({1+2,3})'])
    expect(calc.expression(calls, formula)).toEqual({ kind: 'unsupported', value: '#ERROR!' });
  for (const formula of ['SUMPRODUCT(Data!A:A)', 'SUMPRODUCT(Data!A1:A1000*Data!A1:ZZ1)'])
    expect(calc.expression(calls, formula)).toEqual({ kind: 'unsupported', value: '#LIMIT!' });
  expect(calc(calls, 'A1')).toBe(78);
  const count = calc.diagnostics().evaluations;
  const unrelated = structuredClone(book);
  unrelated.sheets.find((s) => s.name === 'Data')!.cells.F2 = { value: '21' };
  const updated = calc.update(unrelated.sheets);
  expect(
    updated(
      unrelated.sheets.find((s) => s.name === 'Calls')!,
      'A1',
    ),
  ).toBe(78);
  expect(updated.diagnostics().evaluations).toBe(count);
});
it('exports an unchanged shaped-reference workbook byte-identically', async () => {
  const data = input();
  const file = await importFile(
    Object.assign(new File([data], 'Arrays.xlsx'), { arrayBuffer: async () => data }),
  );
  const blob = await exportOffice(file);
  const bytes = await new Promise<ArrayBuffer>((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.readAsArrayBuffer(blob);
  });
  expect(bytes).toEqual(data);
});
function stage(book: WorkbookContent, name: string) {
  const data = book.sheets.find((s) => s.name === 'Data')!;
  if (name === 'edit') data.cells.A2 = { value: '64' };
  if (name === 'delete') delete data.cells.A2;
  if (name === 'error') data.cells.A2 = { value: '=1/0' };
  if (name === 'formula')
    book.sheets.find((s) => s.name === 'Calls')!.cells.D8 = {
      value: '=SUMPRODUCT((Data!A1:A3>=4)*Data!B1:B3)',
    };
  if (name === 'table') data.cells.H3 = { value: '4' };
  return book;
}
it.each([{ stage: 'source', cells: native.initial }, ...native.stages])(
  'matches native array/reference values and types at $stage',
  async (oracle) => {
    const original = await importWorkbook(input());
    const calc = calculator(original.sheets);
    const calls = original.sheets.find((s) => s.name === 'Calls')!;
    for (const item of native.initial) calc(calls, item.ref);
    const updated = stage(structuredClone(original), oracle.stage),
      next = calc.update(updated.sheets);
    const sheet = updated.sheets.find((s) => s.name === 'Calls')!;
    for (const item of oracle.cells)
      expect(next.result(sheet, item.ref), `${oracle.stage} ${item.ref} ${item.formula}`).toEqual({
        kind: item.kind,
        value: item.value,
      });
    const restored = next.update(original.sheets);
    for (const item of native.initial)
      expect(restored.result(calls, item.ref), `restore ${item.ref}`).toEqual({
        kind: item.kind,
        value: item.value,
      });
  },
);
it.each(native.stages)(
  'exports raw native caches and preserves unrelated payloads at $stage',
  async (oracle) => {
    const data = input();
    const file = await importFile(
      Object.assign(new File([data], 'Intersection.xlsx'), { arrayBuffer: async () => data }),
    );
    stage(file.content as WorkbookContent, oracle.stage);
    const blob = await exportOffice(file);
    const bytes = await new Promise<ArrayBuffer>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.readAsArrayBuffer(blob);
    });
    const before = await JSZip.loadAsync(data),
      after = await JSZip.loadAsync(bytes);
    const book = file.content as WorkbookContent,
      calls = book.sheets.find((s) => s.name === 'Calls')!;
    const xml = new DOMParser().parseFromString(
      await after.file(calls.sourcePath!)!.async('string'),
      'application/xml',
    );
    for (const item of oracle.cells) {
      const cell = Array.from(xml.getElementsByTagNameNS('*', 'c')).find(
        (c) => c.getAttribute('r') === item.ref,
      )!;
      expect(cell.getElementsByTagNameNS('*', 'f')[0].textContent).toBe(item.formula.slice(1));
      const type =
        item.kind === 'error'
          ? 'e'
          : typeof item.value === 'boolean'
            ? 'b'
            : typeof item.value === 'string'
              ? 'str'
              : 'n';
      expect(cell.getAttribute('t') || 'n', item.ref).toBe(type);
      expect(cell.getElementsByTagNameNS('*', 'v')[0]?.textContent || '', item.ref).toBe(
        typeof item.value === 'boolean' ? (item.value ? '1' : '0') : String(item.value),
      );
    }
    const allowed = new Set([
      'xl/workbook.xml',
      'xl/calcChain.xml',
      'xl/_rels/workbook.xml.rels',
      '[Content_Types].xml',
      ...book.sheets.map((s) => s.sourcePath),
    ]);
    for (const path of Object.keys(before.files))
      if (!before.files[path].dir && !allowed.has(path))
        expect(await after.file(path)!.async('uint8array'), path).toEqual(
          await before.file(path)!.async('uint8array'),
        );
    const names = (xml: string) =>
      new DOMParser()
        .parseFromString(xml, 'application/xml')
        .getElementsByTagNameNS('*', 'definedNames')[0].outerHTML;
    expect(names(await after.file('xl/workbook.xml')!.async('string'))).toBe(
      names(await before.file('xl/workbook.xml')!.async('string')),
    );
    expect(file.original!.data).toEqual(data);
  },
);
