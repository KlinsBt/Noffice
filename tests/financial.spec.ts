import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
test('loan calculations replace stale caches, update with terms, undo, persist and export', async ({
  page,
}) => {
  const book = new ExcelJS.Workbook(),
    sheet = book.addWorksheet('Loan');
  sheet.getCell('A1').value = 0.06;
  sheet.getCell('A2').value = 12;
  sheet.getCell('A3').value = 1200;
  for (const [ref, formula] of Object.entries({
    B1: 'PMT(A1/12,A2,A3)',
    B2: 'IPMT(A1/12,1,A2,A3)',
    B3: 'PPMT(A1/12,1,A2,A3)',
    C1: 'B1*A2+A3',
  }))
    sheet.getCell(ref).value = { formula, result: 999 };
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Loan.xlsx',
    mimeType: 'application/octet-stream',
    buffer: Buffer.from(await book.xlsx.writeBuffer()),
  });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  await expect(cell('B2')).toHaveText('-6');
  await expect
    .poll(async () => Number(await cell('B1').textContent()))
    .toBeCloseTo(-103.2797156484988, 9);
  await cell('A1').dblclick();
  await page.getByRole('textbox', { name: 'Edit A1', exact: true }).fill('0');
  await page.keyboard.press('Enter');
  await expect(cell('B1')).toHaveText('-100');
  await expect(cell('B2')).toHaveText('0');
  await expect(cell('B3')).toHaveText('-100');
  await expect(cell('C1')).toHaveText('0');
  await page.keyboard.press('Control+z');
  await expect(cell('B2')).toHaveText('-6');
  await page.keyboard.press('Control+y');
  await expect(cell('B2')).toHaveText('0');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Loan', exact: true }).click();
  await expect(cell('B1')).toHaveText('-100');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  for (const [ref, want] of Object.entries({ B1: -100, B2: 0, B3: -100, C1: 0 })) {
    const cellXml = xml.match(new RegExp(`<c\\b[^>]*r="${ref}"[^>]*>([\\s\\S]*?)</c>`))![1];
    expect(Number(cellXml.match(/<v>(.*?)<\/v>/)![1])).toBe(want);
  }
});
