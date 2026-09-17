import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
test('COUNT and COUNTA distinguish errors, blanks and formula-empty cells through edits and exports', async ({
  page,
}) => {
  const book = new ExcelJS.Workbook(),
    sheet = book.addWorksheet('Aggregates');
  sheet.getCell('A1').value = 2;
  sheet.getCell('A2').value = 4;
  sheet.getCell('A3').value = '6';
  sheet.getCell('A4').value = true;
  sheet.getCell('A5').value = false;
  sheet.getCell('A7').value = { formula: '""', result: '' };
  sheet.getCell('A8').value = { formula: 'NA()', result: { error: '#N/A' } };
  for (const [ref, formula] of [
    ['B1', 'COUNT(A1:A8)'],
    ['B2', 'COUNTA(A1:A8)'],
    ['B3', 'SUM(A1:A7)'],
    ['B4', 'SUM(TRUE,"3",2)'],
    ['B5', 'SUM(A1:A8)'],
    ['C1', 'B1*2'],
  ])
    sheet.getCell(ref).value = { formula, result: 100 };
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Aggregates.xlsx',
    mimeType: 'application/octet-stream',
    buffer: Buffer.from(await book.xlsx.writeBuffer()),
  });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  for (const [ref, value] of [
    ['B1', '2'],
    ['B2', '7'],
    ['B3', '6'],
    ['B4', '6'],
    ['B5', '#N/A'],
  ])
    await expect(cell(ref)).toHaveText(value);
  await cell('A8').dblclick();
  await page.getByRole('textbox', { name: 'Edit A8', exact: true }).fill('10');
  await page.keyboard.press('Enter');
  await expect(cell('B1')).toHaveText('3');
  await expect(cell('B2')).toHaveText('7');
  await expect(cell('B5')).toHaveText('16');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Aggregates', exact: true }).click();
  await expect(cell('B1')).toHaveText('3');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!)),
    xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  for (const [ref, value] of [
    ['B1', '3'],
    ['B2', '7'],
    ['B5', '16'],
    ['C1', '6'],
  ])
    expect(xml.match(new RegExp(`<c\\b[^>]*r="${ref}"[^>]*>([\\s\\S]*?)</c>`))?.[1]).toContain(
      `<v>${value}</v>`,
    );
});
