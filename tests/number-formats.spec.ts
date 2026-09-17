import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import fs from 'node:fs/promises';

test('custom number formats preview, apply to a range, undo, persist and export without changing values', async ({
  page,
}) => {
  const book = new ExcelJS.Workbook(),
    sheet = book.addWorksheet('Amounts');
  sheet.getCell('A1').value = 7.6;
  sheet.getCell('A2').value = -7.6;
  sheet.getCell('B1').value = 0.00035;
  sheet.getCell('C1').value = { formula: 'TEXT(B1,"#,##0.00%")', result: '0.04%' };
  sheet.getCell('D1').value = 5.25;
  sheet.getCell('E1').value = 0.1702084490740741;
  sheet.getCell('E1').numFmt = 'h:m:s.00 a/p';
  sheet.pageSetup = { printArea: 'A1:E2', orientation: 'landscape' };
  const input = Buffer.from(await book.xlsx.writeBuffer());
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Excel', exact: true }).waitFor();
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Formatting.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: input,
  });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  await expect(cell('C1')).toHaveText('0.04%');
  await expect(cell('E1')).toHaveText('4:5:6.01 a');
  await cell('A1').click();
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Control+1');
  const modal = page.getByRole('dialog', { name: 'Format cells' });
  await expect(modal).toBeVisible();
  await page.getByRole('textbox', { name: 'Format code' }).fill('£000.00');
  await expect(page.getByLabel('Format preview')).toHaveText('-£007.60');
  // Enter in the dialog must not start editing a worksheet cell.
  await page.getByRole('button', { name: 'Apply format' }).focus();
  await page.keyboard.press('Enter');
  await expect(modal).toHaveCount(0);
  await expect(cell('A1')).toHaveText('£007.60');
  await expect(cell('A2')).toHaveText('-£007.60');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('A1')).toHaveText('7.6');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(cell('A1')).toHaveText('£007.60');
  await cell('D1').click();
  await page.getByRole('button', { name: 'Format cells', exact: true }).click();
  await page.getByRole('combobox', { name: 'Format category' }).selectOption({ label: 'Fraction' });
  await expect(page.getByLabel('Format preview')).toHaveText('5  1/4 ');
  await page.getByRole('button', { name: 'Apply format' }).click();
  await expect(cell('D1')).toHaveText('5  1/4 ');
  await page.getByRole('button', { name: 'Format cells', exact: true }).click();
  await page.getByRole('textbox', { name: 'Format code' }).fill('0.' + '0'.repeat(101));
  await expect(page.getByRole('button', { name: 'Apply format' })).toBeDisabled();
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Formatting', exact: true }).click();
  await expect(cell('A1')).toHaveText('£007.60');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  const exported = new ExcelJS.Workbook();
  await exported.xlsx.load(Uint8Array.from(output).buffer);
  expect(exported.worksheets[0].getCell('A1').value).toBe(7.6);
  expect(exported.worksheets[0].getCell('A2').value).toBe(-7.6);
  expect(exported.worksheets[0].getCell('A1').numFmt).toBe('£000.00');
  expect(exported.worksheets[0].getCell('D1').numFmt).toBe('# ??/??');
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && !['xl/styles.xml', 'xl/worksheets/sheet1.xml'].includes(path))
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await page.screenshot({ path: 'test-results/excel-number-formats.png' });
});
