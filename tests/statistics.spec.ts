import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';

test('statistical formulas update after edits and undo, persist and export recalculated caches', async ({
  page,
}) => {
  const book = new ExcelJS.Workbook(),
    sheet = book.addWorksheet('Statistics');
  [2, 4, 4, 8].forEach((v, i) => (sheet.getCell(`A${i + 1}`).value = v));
  sheet.getCell('B1').value = { formula: 'VAR.S(A1:A4)', result: 0 };
  sheet.getCell('B2').value = { formula: 'PERCENTILE.INC(A1:A4,0.5)', result: 0 };
  sheet.getCell('B3').value = { formula: 'RANK.AVG(A2,A1:A4)', result: 0 };
  sheet.getCell('C1').value = { formula: 'B1*2', result: 0 };
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Excel', exact: true }).waitFor();
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Statistics.xlsx',
    mimeType: 'application/octet-stream',
    buffer: Buffer.from(await book.xlsx.writeBuffer()),
  });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  await expect(cell('B2')).toHaveText('4');
  await expect(cell('B3')).toHaveText('2.5');
  await cell('A4').dblclick();
  await page.getByRole('textbox', { name: 'Edit A4', exact: true }).fill('6');
  await page.keyboard.press('Enter');
  await expect.poll(async () => Number(await cell('B1').textContent())).toBeCloseTo(8 / 3, 9);
  await page.keyboard.press('Control+z');
  await expect(cell('A4')).toHaveText('8');
  await page.keyboard.press('Control+y');
  await expect(cell('A4')).toHaveText('6');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Statistics', exact: true }).click();
  await expect.poll(async () => Number(await cell('B1').textContent())).toBeCloseTo(8 / 3, 9);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!)),
    xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  const value = (ref: string) =>
    Number(
      xml
        .match(new RegExp(`<c\\b[^>]*r="${ref}"[^>]*>([\\s\\S]*?)</c>`))?.[1]
        .match(/<v>(.*?)<\/v>/)?.[1],
    );
  expect(value('B1')).toBeCloseTo(8 / 3, 12);
  expect(value('C1')).toBeCloseTo(16 / 3, 12);
});
