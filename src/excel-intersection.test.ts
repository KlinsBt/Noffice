import { expect, it } from 'vitest';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import native from '../tests/fixtures/native-excel-intersection.json';
import { calculator } from './formulas';
import { importWorkbook } from './xlsx-import';
import { importFile, exportOffice } from './formats';
import type { WorkbookContent } from './model';
const input = () => {
  const bytes = fs.readFileSync('tests/fixtures/excel-intersection.xlsx');
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(native.sourceSha256);
  return Uint8Array.from(bytes).buffer;
};
function stage(book: WorkbookContent, name: string) {
  const data = book.sheets.find((s) => s.name === 'Data')!;
  if (name === 'edit') data.cells.A2 = { value: '64' };
  if (name === 'delete') delete data.cells.A2;
  if (name === 'error') data.cells.A2 = { value: '=1/0' };
  return book;
}
it.each([{ stage: 'source', cells: native.initial }, ...native.stages])(
  'matches native scalar/reference values and types at $stage',
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
it('reads only the intersecting precedent for scalar full-column references', async () => {
  const book = await importWorkbook(input()),
    calc = calculator(book.sheets);
  const calls = book.sheets.find((s) => s.name === 'Calls')!;
  expect(calc.result(calls, 'E4')).toEqual({ value: 16, kind: 'value' });
  const evaluated = calc.diagnostics().evaluations;
  const edited = structuredClone(book);
  edited.sheets.find((s) => s.name === 'Data')!.cells.A8 = { value: '=NA()' };
  const next = calc.update(edited.sheets);
  expect(
    next(
      edited.sheets.find((s) => s.name === 'Calls')!,
      'E4',
    ),
  ).toBe(16);
  expect(next.diagnostics().evaluations).toBe(evaluated);
  expect(calc.expression(calls, 'Data!A1:A5')).toEqual({
    value: '#UNSUPPORTED!',
    kind: 'unsupported',
  });
});
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
