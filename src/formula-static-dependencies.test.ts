import { expect, it } from 'vitest';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { bindFormulaReferences, calculator } from './formulas';
import { formulaCellIdentity as key, staticFormulaDependents } from './formula-static-dependencies';
import type { Sheet, WorkbookContent } from './model';
import { importWorkbook } from './xlsx-import';
import { importFile, exportOffice } from './formats';
import native from '../tests/fixtures/native-excel-static-references.json';

const sheet = (id: string, values: Record<string, string> = {}): Sheet => ({
  id,
  name: id,
  rows: 100,
  cols: 26,
  cells: Object.fromEntries(Object.entries(values).map(([ref, value]) => [ref, { value }])),
});
const bytes = () =>
  Uint8Array.from(fs.readFileSync('tests/fixtures/excel-static-references.xlsx')).buffer;
const nativeValue = (cell: { value: number | string | boolean | null; valueType: string }) =>
  cell.valueType === 'System.Int32' ? '#DIV/0!' : (cell.value ?? '');

it('retains absolute/mixed axes, source endpoint order and stable sheet identities', () => {
  const a = sheet('data-id'),
    b = { ...sheet('cost-id'), name: "Cost's data" };
  const result = bindFormulaReferences([a, b], a, "=$a1+A$2+SUM('Cost''s data'!$B2:A$1)");
  expect(result.complete).toBe(true);
  expect(result.references).toEqual([
    {
      sheetId: a.id,
      kind: 'cell',
      from: { row: 0, col: 0, absoluteRow: false, absoluteCol: true },
      to: { row: 0, col: 0, absoluteRow: false, absoluteCol: true },
    },
    {
      sheetId: a.id,
      kind: 'cell',
      from: { row: 1, col: 0, absoluteRow: true, absoluteCol: false },
      to: { row: 1, col: 0, absoluteRow: true, absoluteCol: false },
    },
    {
      sheetId: b.id,
      kind: 'range',
      from: { row: 1, col: 1, absoluteRow: false, absoluteCol: true },
      to: { row: 0, col: 0, absoluteRow: true, absoluteCol: false },
    },
  ]);
  const renamed = { ...b, name: 'Renamed' };
  expect(bindFormulaReferences([a, renamed], a, '=Renamed!$B2:A$1').references[0]).toEqual(
    result.references[2],
  );
  expect(key('a:b', 'C1')).not.toBe(key('a', 'b:C1'));
});
it('represents entire rows/columns as two bounded rectangles with their anchors', () => {
  const s = sheet('Data');
  const { references, complete } = bindFormulaReferences([s], s, '=SUM($A:B)+SUM($2:3)');
  expect(complete).toBe(true);
  expect(references).toHaveLength(2);
  expect(references[0]).toMatchObject({
    kind: 'columns',
    from: { col: 0, absoluteCol: true },
    to: { col: 1, row: 1048575 },
  });
  expect(references[1]).toMatchObject({
    kind: 'rows',
    from: { row: 1, absoluteRow: true },
    to: { row: 2, col: 16383 },
  });
});
it('distinguishes incomplete or failed bindings from an exhaustive empty result', () => {
  const s = sheet('Data');
  for (const formula of [
    '=Rate+A1',
    '=SUM(Table1[Amount])',
    '=INDIRECT("A1")',
    '=SUBTOTAL(9,A1:A3)',
    "='[other.xlsx]Data'!A1",
    '=Missing!A1',
    '=SUM($XFE$1:A1)',
    '=IF(TRUE,A1,#REF!)',
    '=SUM(',
  ])
    expect(bindFormulaReferences([s], s, formula).complete, formula).toBe(false);
  expect(bindFormulaReferences([s], s, '="A1"+42')).toEqual({ complete: true, references: [] });
});
it('discovers inactive branches without evaluating them or changing observed cycles/errors', () => {
  const s = sheet('Data', {
    A1: '=IF(TRUE,4,B1)',
    B1: '=A1',
    C1: '=IF(TRUE,3,1/0)',
    D1: '=IF(FALSE,B1,42)',
  });
  const calc = calculator([s]);
  expect(bindFormulaReferences([s], s, s.cells.A1.value).references[0].from.row).toBe(0);
  expect(calc(s, 'A1')).toBe(4);
  const before = calc.diagnostics();
  const dirty = staticFormulaDependents([s], [key(s.id, 'B1')]);
  expect(dirty).toEqual(new Set(['A1', 'B1', 'D1'].map((ref) => key(s.id, ref))));
  expect(calc.diagnostics()).toEqual(before);
  expect(calc(s, 'C1')).toBe(3);
});
it('finds transitive rectangle consumers, blank precedents and reversed ranges', () => {
  const a = sheet('Data', {
    A1: '2',
    C1: '=SUM(B2:A1)',
    C2: '=SUM(A:A)',
    C3: '=SUM(1:2)',
    D1: '=99',
  });
  const b = sheet('Summary', { A1: '=Data!C1' });
  const dirty = staticFormulaDependents([a, b], [key(a.id, 'A2')]);
  for (const ref of ['A2', 'C1', 'C2', 'C3']) expect(dirty.has(key(a.id, ref))).toBe(true);
  expect(dirty.has(key(b.id, 'A1'))).toBe(true);
  expect(dirty.has(key(a.id, 'D1'))).toBe(false);
});
it('propagates conservative unknowns, removes deleted formulas and falls back on either work budget', () => {
  const s = sheet('Data', { A1: '1', B1: '=Rate', C1: '=B1', D1: '=99', E1: '=A1', F1: '=A1' });
  expect(staticFormulaDependents([s], [])).toEqual(
    new Set(['B1', 'C1'].map((ref) => key(s.id, ref))),
  );
  expect(staticFormulaDependents([s], [key(s.id, 'A1')], 1).has(key(s.id, 'D1'))).toBe(true);
  expect(staticFormulaDependents([s], [key(s.id, 'A1')], 3).has(key(s.id, 'D1'))).toBe(true);
  const deleted = { ...s, cells: { A1: s.cells.A1, B1: { value: '7' }, C1: s.cells.C1 } };
  expect(staticFormulaDependents([deleted], []).size).toBe(0);
  expect(staticFormulaDependents([deleted], [key(s.id, 'B1')])).toEqual(
    new Set(['B1', 'C1'].map((ref) => key(s.id, ref))),
  );
});
it('matches the pinned native input and all three edit/delete/inactive-branch stages', async () => {
  expect(createHash('sha256').update(new Uint8Array(bytes())).digest('hex')).toBe(
    native.sourceHash,
  );
  expect(native.conditionalPrecedents).toBe('A1,A2:B2');
  for (const stage of [{ stage: 'initial', cells: native.initial }, ...native.stages]) {
    const book = await importWorkbook(bytes()),
      data = book.sheets.find((s) => s.name === 'Data')!;
    if (stage.stage === 'edit') data.cells.A1.value = '5';
    if (stage.stage === 'delete') delete data.cells.A1;
    if (stage.stage === 'branch') data.cells.B2.value = '9';
    const calc = calculator(book.sheets);
    for (const cell of stage.cells) {
      const s = book.sheets.find((s) => s.name === cell.sheet)!;
      expect(calc(s, cell.address), `${stage.stage} ${cell.sheet}!${cell.address}`).toBe(
        nativeValue(cell),
      );
    }
  }
});
it.each(['edit', 'delete', 'branch'])(
  'writes native %s caches while preserving formula anchors and original bytes',
  async (stage) => {
    const data = bytes(),
      hash = createHash('sha256').update(new Uint8Array(data)).digest('hex');
    const file = await importFile(
      Object.assign(new File([data], 'References.xlsx'), { arrayBuffer: async () => data }),
    );
    const book = file.content as WorkbookContent,
      s = book.sheets.find((s) => s.name === 'Data')!;
    if (stage === 'edit') s.cells.A1.value = '5';
    if (stage === 'delete') delete s.cells.A1;
    if (stage === 'branch') s.cells.B2.value = '9';
    const blob = await exportOffice(file);
    const output = await new Promise<ArrayBuffer>((resolve) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result as ArrayBuffer);
      r.readAsArrayBuffer(blob);
    });
    const zip = await JSZip.loadAsync(output);
    const expected = native.stages.find((s) => s.stage === stage)!;
    for (const cell of expected.cells.filter((c) => c.formula.startsWith('='))) {
      const sheet = book.sheets.find((s) => s.name === cell.sheet)!;
      const xml = new DOMParser().parseFromString(
        await zip.file(sheet.sourcePath!)!.async('string'),
        'application/xml',
      );
      const node = Array.from(xml.getElementsByTagNameNS('*', 'c')).find(
        (c) => c.getAttribute('r') === cell.address,
      )!;
      expect(node.getElementsByTagNameNS('*', 'f')[0].textContent).toBe(cell.formula.slice(1));
      expect(node.getElementsByTagNameNS('*', 'v')[0].textContent, `${stage} ${cell.address}`).toBe(
        typeof cell.value === 'boolean' ? (cell.value ? '1' : '0') : String(nativeValue(cell)),
      );
      if (cell.valueType === 'System.Int32') expect(node.getAttribute('t')).toBe('e');
    }
    expect(createHash('sha256').update(new Uint8Array(data)).digest('hex')).toBe(hash);
  },
);
