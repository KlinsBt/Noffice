import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import fs from 'node:fs/promises';
test('filter values, edit, reapply, combine columns, undo, persist, export and clear', async ({
  page,
}) => {
  const book = new ExcelJS.Workbook(),
    ws = book.addWorksheet('Data');
  ws.addRows([
    ['Amount', 'Group'],
    [2, 'keep'],
    [4, 'drop'],
    [6, 'keep'],
  ]);
  ws.getCell('D1').value = { formula: 'SUBTOTAL(9,A2:A4)', result: 12 };
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Excel', exact: true }).waitFor();
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles({
      name: 'Filters.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: Buffer.from(await book.xlsx.writeBuffer()),
    });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  await cell('A1').click();
  await cell('B4').click({ modifiers: ['Shift'] });
  await page.getByRole('button', { name: 'Filter range', exact: true }).click();
  await page.getByRole('combobox', { name: 'Filter column' }).selectOption({ label: 'Group' });
  await page.getByRole('checkbox', { name: 'drop', exact: true }).uncheck();
  await page.getByRole('button', { name: 'Apply filter', exact: true }).click();
  await expect(cell('A3')).toHaveCount(0);
  await expect(cell('D1')).toHaveText('8');
  await cell('B4').dblclick();
  await page.keyboard.press('Control+A');
  await page.keyboard.type('drop');
  await page.keyboard.press('Enter');
  await expect(cell('A4')).toHaveText('6');
  await expect(cell('D1')).toHaveText('8');
  await page.getByRole('button', { name: 'Reapply filters', exact: true }).click();
  await expect(cell('A4')).toHaveCount(0);
  await expect(cell('D1')).toHaveText('2');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('D1')).toHaveText('8');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(cell('D1')).toHaveText('2');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Filters', exact: true }).click();
  await expect(cell('D1')).toHaveText('2');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.mkdir('.local/filter-validation', { recursive: true });
  await fs.writeFile('.local/filter-validation/browser-export.xlsx', output);
  const exported = new ExcelJS.Workbook();
  await exported.xlsx.load(Uint8Array.from(output).buffer);
  expect(exported.worksheets[0].getRow(4).hidden).toBe(true);
  expect(exported.worksheets[0].getCell('D1').result).toBe(2);
  await page.getByRole('button', { name: 'Clear filters', exact: true }).click();
  await expect(cell('A3')).toHaveText('4');
  await expect(cell('A4')).toHaveText('6');
  await expect(cell('D1')).toHaveText('12');
  await cell('A1').click();
  await page.getByRole('button', { name: 'Filter range', exact: true }).click();
  await page.getByRole('combobox', { name: 'Filter match' }).selectOption('greaterThan');
  await page.getByRole('textbox', { name: 'Filter value', exact: true }).fill('3');
  await page.getByRole('button', { name: 'Apply filter', exact: true }).click();
  await expect(cell('A2')).toHaveCount(0);
  await expect(cell('D1')).toHaveText('10');
});
