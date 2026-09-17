import { it, expect } from 'vitest';
import fs from 'node:fs';
import JSZip from 'jszip';
import { calculator } from './formulas';
import { importWorkbook } from './xlsx-import';
import { importFile, exportOffice } from './formats';
import { enterArrayFormula, arrayEditError } from './sheet-arrays';
import { restoreTableMetadata } from './sheet-tables';
import native from '../tests/fixtures/native-excel-array-contexts.json';
const bytes = () =>
  Uint8Array.from(fs.readFileSync('tests/fixtures/excel-array-contexts.xlsx')).buffer;
const input = () => importWorkbook(bytes());
it('matches all 96 native ordinary/CSE snapshots through precedent revisions and recovery', async () => {
  const original = await input();
  const calc = calculator(original.sheets);
  for (const stage of [...native.stages, native.stages[0]]) {
    const book = structuredClone(original);
    const data = book.sheets.find((s) => s.name === 'Data')!;
    if (stage.stage === 'edit') data.cells.A2 = { value: '0' };
    if (stage.stage === 'error') data.cells.A2 = { value: '=1/0' };
    const calls = book.sheets.find((s) => s.name === 'Calls')!;
    calc.update(book.sheets);
    for (const expected of stage.cells)
      expect(calc.result(calls, expected.ref), `${stage.stage}:${expected.ref}`).toEqual({
        value: expected.value,
        kind: expected.kind,
      });
  }
});
it('changes entry intent without changing formula text, preserving history snapshots and lazy scalar branches', async () => {
  const source = await input(),
    calls = source.sheets.find((s) => s.name === 'Calls')!;
  const entered = enterArrayFormula(source, calls.id, 'B1', calls.cells.B1.value, true);
  expect(
    calculator(entered.sheets)(
      entered.sheets.find((s) => s.id === calls.id)!,
      'B1',
    ),
  ).toBe(13);
  expect(calculator(source.sheets)(calls, 'B1')).toBe(0);
  const restored = enterArrayFormula(entered, calls.id, 'B1', calls.cells.B1.value, false);
  expect(
    calculator(restored.sheets)(
      restored.sheets.find((s) => s.id === calls.id)!,
      'B1',
    ),
  ).toBe(0);
  const lazy = enterArrayFormula(
    source,
    calls.id,
    'D1',
    '=IF(FALSE,SUMPRODUCT(Data!A:A),42)',
    true,
  );
  expect(
    calculator(lazy.sheets)(
      lazy.sheets.find((s) => s.id === calls.id)!,
      'D1',
    ),
  ).toBe(42);
  const huge = enterArrayFormula(source, calls.id, 'D1', '=SUMPRODUCT(Data!A:A)', true);
  expect(
    calculator(huge.sheets).result(
      huge.sheets.find((s) => s.id === calls.id)!,
      'D1',
    ).value,
  ).toBe('#LIMIT!');
  expect(() => enterArrayFormula(source, calls.id, 'D1', 'text', true)).toThrow('Enter a formula');
});
it('restores missing legacy array identities while retaining edits, and rejects ambiguous source formulas', async () => {
  const source = await input(),
    old = structuredClone(source);
  for (const s of old.sheets) delete s.arrayFormulas;
  old.sheets.find((s) => s.name === 'Data')!.cells.A2 = { value: '0' };
  const restored = restoreTableMetadata(old, source),
    calls = restored.sheets.find((s) => s.name === 'Calls')!;
  expect(calls.arrayFormulas).toHaveLength(16);
  expect(calculator(restored.sheets)(calls, 'C1')).toBe(8);
  old.sheets.find((s) => s.name === 'Calls')!.cells.C1.value = '=99';
  expect(() => restoreTableMetadata(old, source)).toThrow('Saved array formulas differ');
});
it('guards multi-cell array members before edits and never returns stale follower caches as recalculated values', async () => {
  const book = await input(),
    calls = book.sheets.find((s) => s.name === 'Calls')!;
  calls.arrayFormulas = [{ anchor: 'C1', ref: 'C1:C2' }];
  const changed = structuredClone(calls);
  changed.cells.C2.value = '123';
  expect(arrayEditError(calls, changed)).toContain('multi-cell array');
  expect(calculator(book.sheets).result(calls, 'C2').value).toBe('#UNSUPPORTED!');
  expect(() => enterArrayFormula(book, calls.id, 'C1', '=3', true)).toThrow('multi-cell array');
  const legacy = structuredClone(book);
  const saved = legacy.sheets.find((s) => s.id === calls.id)!;
  delete saved.arrayFormulas;
  saved.cells.C2.value = '123';
  expect(() => restoreTableMetadata(legacy, book)).toThrow('Saved array formulas differ');
});
it('exports CSE identities, actual recalculated caches and untouched package payloads', async () => {
  const data = bytes();
  const file = await importFile(
    Object.assign(new File([data], 'Arrays.xlsx'), { arrayBuffer: async () => data }),
  );
  if (file.content.kind !== 'excel') throw Error('Wrong fixture');
  const calls = file.content.sheets.find((s) => s.name === 'Calls')!;
  const next = enterArrayFormula(file.content, calls.id, 'B1', calls.cells.B1.value, true);
  const blob = await exportOffice({ ...file, content: next });
  const buffer = await new Promise<ArrayBuffer>((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as ArrayBuffer);
    r.onerror = reject;
    r.readAsArrayBuffer(blob);
  });
  const zip = await JSZip.loadAsync(buffer),
    original = await JSZip.loadAsync(data);
  const xml = new DOMParser().parseFromString(
    await zip.file('xl/worksheets/sheet2.xml')!.async('string'),
    'application/xml',
  );
  const b1 = [...xml.getElementsByTagName('c')].find((c) => c.getAttribute('r') === 'B1')!;
  expect(b1.getElementsByTagName('f')[0].getAttribute('t')).toBe('array');
  expect(b1.getElementsByTagName('v')[0].textContent).toBe('13');
  expect(await zip.file('xl/theme/theme1.xml')!.async('string')).toBe(
    await original.file('xl/theme/theme1.xml')!.async('string'),
  );
  const reopened = await importWorkbook(buffer);
  const sheet = reopened.sheets.find((s) => s.name === 'Calls')!;
  expect(sheet.arrayFormulas).toContainEqual({ anchor: 'B1', ref: 'B1' });
  expect(calculator(reopened.sheets)(sheet, 'B1')).toBe(13);
});
