import { it, expect } from 'vitest';
import JSZip from 'jszip';
import { formulaWorkbook } from '../tests/fixtures/excel-formula-workbook';
import { contentSchema, newFile } from './model';
import { importWorkbook, elements, parseXML } from './xlsx-import';
import { exportRetainedWorkbook } from './xlsx-preserve';
import { calculator } from './formulas';

async function bytes(blob: Blob) {
  return new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = reject;
    reader.readAsArrayBuffer(blob);
  });
}
it('imports table metadata, saves it in native backups and updates indirect/named/table dependent caches without changing other package parts', async () => {
  const input = await formulaWorkbook();
  const content = await importWorkbook(input);
  const file = newFile('excel', 'Formula workbook', content);
  file.original = { name: 'Formula workbook.xlsx', data: input };
  expect(contentSchema.parse(content)).toEqual(content);
  const sheet = content.sheets[0];
  expect(sheet.tables?.[0]).toMatchObject({
    name: 'SalesData',
    columns: ['Team', 'Amount', 'Units', 'Label'],
  });
  expect(await bytes(await exportRetainedWorkbook(file))).toEqual(input);
  sheet.cells.B2.value = '50';
  const output = await bytes(await exportRetainedWorkbook(file));
  const a = await JSZip.loadAsync(input),
    b = await JSZip.loadAsync(output);
  expect(Object.keys(b.files).sort()).toEqual(Object.keys(a.files).sort());
  for (const name of Object.keys(a.files))
    if (!a.files[name].dir && !['xl/workbook.xml', sheet.sourcePath].includes(name))
      expect(await b.file(name)!.async('uint8array'), name).toEqual(
        await a.file(name)!.async('uint8array'),
      );
  const xml = parseXML(await b.file(sheet.sourcePath!)!.async('string'));
  const cached = (ref: string) =>
    elements(
      elements(xml, 'c').find((e) => e.getAttribute('r') === ref)!,
      'v',
    )[0]?.textContent;
  expect(cached('J2')).toBe('100');
  expect(cached('J11')).toBe('140');
  expect(cached('J12')).toBe('168');
  expect(cached('J14')).toBe('140');
  expect(cached('J16')).toBe('140');
  expect(cached('L2')).toBe('336');
  expect(cached('L3')).toBe('987');
  const again = await importWorkbook(output);
  expect(calculator(again.sheets)(again.sheets[0], 'L2')).toBe(336);
});
