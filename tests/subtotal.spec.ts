import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import fs from 'node:fs/promises';

test('row hiding updates visible subtotals, supports undo/reload and exports recalculated caches', async ({
  page,
}) => {
  const book = new ExcelJS.Workbook(),
    ws = book.addWorksheet('Data');
  ws.getCell('A1').value = 'Amount';
  ws.getCell('A2').value = 2;
  ws.getCell('A3').value = 4;
  ws.getCell('A4').value = 6;
  ws.getCell('B1').value = { formula: 'SUBTOTAL(9,A2:A4)', result: 12 };
  ws.getCell('C1').value = { formula: 'SUBTOTAL(109,A2:A4)', result: 12 };
  ws.getCell('D1').value = { formula: 'C1*2', result: 24 };
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Excel', exact: true }).waitFor();
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles({
      name: 'Subtotals.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: Buffer.from(await book.xlsx.writeBuffer()),
    });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  await expect(cell('C1')).toHaveText('12');
  await cell('A3').click();
  await page.getByRole('button', { name: 'Sheet settings', exact: true }).click();
  await page.getByRole('button', { name: 'Hide selected rows', exact: true }).click();
  await expect(cell('A3')).toHaveCount(0);
  await expect(cell('B1')).toHaveText('12');
  await expect(cell('C1')).toHaveText('8');
  await expect(cell('D1')).toHaveText('16');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('A3')).toHaveText('4');
  await expect(cell('C1')).toHaveText('12');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(cell('C1')).toHaveText('8');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Subtotals', exact: true }).click();
  await expect(cell('C1')).toHaveText('8');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  const exported = new ExcelJS.Workbook();
  await exported.xlsx.load(Uint8Array.from(output).buffer);
  expect(exported.worksheets[0].getRow(3).hidden).toBe(true);
  expect(exported.worksheets[0].getCell('C1').result).toBe(8);
  expect(exported.worksheets[0].getCell('D1').result).toBe(16);
  await page.getByRole('button', { name: 'Sheet settings', exact: true }).click();
  await page.getByRole('button', { name: 'Unhide all rows', exact: true }).click();
  await expect(cell('A3')).toHaveText('4');
  await expect(cell('C1')).toHaveText('12');
});
