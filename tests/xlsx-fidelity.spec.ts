import { test, expect, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import fs from 'node:fs/promises';

async function exportXlsx(page: Page) {
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  return fs.readFile((await (await pending).path())!);
}
async function upload(page: Page, buffer: Buffer) {
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Excel', exact: true }).waitFor();
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Authored.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer,
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Authored',
  );
}
async function source() {
  const book = new ExcelJS.Workbook(),
    s = book.addWorksheet('Price form');
  s.columns = [{ width: 31 }, { width: 15 }, { width: 24 }, { width: 10 }];
  s.getRow(1).height = 48;
  s.getRow(6).hidden = true;
  s.getCell('A6').value = 'Hidden content';
  s.mergeCells('A1:D1');
  s.getCell('A1').value = 'Project price schedule';
  s.getCell('A1').font = { name: 'Arial', size: 16, bold: true, color: { theme: 0 } };
  s.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { theme: 4 } };
  s.getCell('A1').alignment = { wrapText: true, vertical: 'middle' };
  s.getCell('A3').value = '001234';
  s.getCell('B3').value = 2;
  s.getCell('B3').protection = { locked: false };
  s.getCell('C3').value = 75;
  s.getCell('C3').numFmt = '#,##0.00 "EUR"';
  s.getCell('D3').value = { formula: 'B3*C3', result: 150 };
  s.getCell('A4').value = new Date('2025-12-03T00:00:00Z');
  s.getCell('A4').numFmt = 'dd.mm.yyyy';
  s.getCell('B3').note = 'Enter quantity';
  s.getCell('B3').dataValidation = { type: 'whole', operator: 'greaterThan', formulae: [0] };
  s.views = [{ state: 'frozen', ySplit: 2, showGridLines: false }];
  s.pageSetup = { orientation: 'landscape', paperSize: 9, printArea: 'A1:D10' };
  const lookup = book.addWorksheet('Lookup', { state: 'hidden' });
  lookup.getCell('A1').value = 20;
  lookup.getCell('B1').value = 'Matched expert';
  s.getCell('A5').value = {
    formula: 'IFERROR(VLOOKUP(20,Lookup!A1:B1,2,FALSE),"wrong")',
    result: 'Matched expert',
  };
  book.definedNames.add('Lookup!$B$1', 'ExpertChoices');
  s.getCell('B5').dataValidation = {
    type: 'list',
    formulae: ['ExpertChoices'],
    showErrorMessage: true,
  };
  s.getCell('B5').protection = { locked: false };
  s.getCell('B5').value = 'Matched expert';
  await s.protect('test-only', { spinCount: 1 });
  return Buffer.from(await book.xlsx.writeBuffer());
}
test('authored XLSX renders its form layout and survives unchanged and edited browser round trips', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  const original = await source();
  await upload(page, original);
  const cell = (name: string) => page.getByRole('gridcell', { name, exact: true });
  await expect(cell('A1')).toHaveAttribute('colspan', '4');
  await expect(cell('A1')).toHaveCSS('height', '64px');
  await expect(cell('A1')).toHaveCSS('color', 'rgb(255, 255, 255)');
  await expect(cell('A3')).toHaveText('001234');
  await expect(cell('A4')).toHaveText('03.12.2025');
  await expect(cell('A5')).toHaveText('Matched expert');
  await expect(
    page.locator('.sheet-tabs').getByRole('button', { name: 'Lookup', exact: true }),
  ).toHaveCount(0);
  await expect(cell('A6')).toHaveCount(0);
  await expect(cell('A1')).toHaveCSS('position', 'sticky');
  expect(await exportXlsx(page)).toEqual(original);
  await cell('B5').click();
  await expect(page.getByRole('combobox', { name: 'Cell choices' })).toContainText(
    'Matched expert',
  );
  await cell('A3').dblclick();
  await expect(page.getByRole('textbox', { name: 'Edit A3', exact: true })).toHaveCount(0);
  await cell('B3').dblclick();
  await page.getByRole('textbox', { name: 'Edit B3', exact: true }).fill('4');
  await page.keyboard.press('Enter');
  await expect(cell('D3')).toHaveText('300');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const edited = await exportXlsx(page),
    a = await JSZip.loadAsync(original),
    b = await JSZip.loadAsync(edited);
  for (const name of Object.keys(a.files))
    if (!a.files[name].dir && !['xl/worksheets/sheet1.xml', 'xl/workbook.xml'].includes(name))
      expect(await b.file(name)!.async('uint8array'), name).toEqual(
        await a.file(name)!.async('uint8array'),
      );
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(Uint8Array.from(edited).buffer);
  expect(book.worksheets[0].getCell('D3').result).toBe(300);
  expect(book.worksheets[0].getCell('A4').value).toBeInstanceOf(Date);
  await page.screenshot({ path: 'test-results/xlsx-preserved-form.png' });
  expect(errors).toEqual([]);
});

test('private supplied workbook retains original bytes and recalculates its supported lookup cells', async ({
  page,
}) => {
  test.skip(
    !process.env.NOFFICE_XLSX_FIXTURE,
    'Set NOFFICE_XLSX_FIXTURE to a local original workbook; private fixtures are not distributed.',
  );
  test.setTimeout(90000);
  const original = await fs.readFile(process.env.NOFFICE_XLSX_FIXTURE!);
  await upload(page, original);
  expect(await exportXlsx(page)).toEqual(original);
  await expect(page.locator('.sheet-tabs>button').filter({ hasText: 'opt. services' })).toHaveCount(
    0,
  );
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveAttribute(
    'colspan',
    '6',
  );
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(Uint8Array.from(original).buffer);
  const sheet = book.worksheets[0];
  for (const ref of ['B27', 'B28']) {
    const expected = String(sheet.getCell(ref).result);
    const cell = page.getByRole('gridcell', { name: ref, exact: true });
    await expect(cell).toHaveText(expected);
  }
  const quantity = page.getByRole('gridcell', { name: 'D27', exact: true });
  await page.getByRole('gridcell', { name: 'A27', exact: true }).click();
  await expect(page.getByRole('combobox', { name: 'Cell choices' })).toBeVisible();
  await quantity.dblclick();
  await page
    .getByRole('textbox', { name: 'Edit D27', exact: true })
    .fill(String(Number(sheet.getCell('D27').value) + 1));
  await page.keyboard.press('Enter');
  const edited = await exportXlsx(page),
    before = await JSZip.loadAsync(original),
    after = await JSZip.loadAsync(edited);
  expect(Object.keys(after.files).sort()).toEqual(Object.keys(before.files).sort());
  for (const name of Object.keys(before.files))
    if (
      !before.files[name].dir &&
      !/^xl\/worksheets\/sheet\d+\.xml$/.test(name) &&
      name !== 'xl/workbook.xml'
    )
      expect(await after.file(name)!.async('uint8array'), name).toEqual(
        await before.file(name)!.async('uint8array'),
      );
  const updated = new ExcelJS.Workbook();
  await updated.xlsx.load(Uint8Array.from(edited).buffer);
  expect(updated.worksheets[0].getCell('D27').value).toBe(Number(sheet.getCell('D27').value) + 1);
  expect(updated.worksheets[0].getCell('F27').result).toBe(
    (Number(sheet.getCell('D27').value) + 1) * Number(sheet.getCell('E27').value),
  );
  await fs.writeFile('test-results/private-price-schedule-edited.xlsx', edited);
  await page.screenshot({ path: 'test-results/private-price-schedule.png' });
});
