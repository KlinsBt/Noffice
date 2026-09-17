import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { formulaWorkbook } from './fixtures/excel-formula-workbook';

test('table and named formulas recalculate after a browser edit, reload and XLSX export', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Excel', exact: true }).waitFor();
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Calculations.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(await formulaWorkbook()),
  });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  await expect(cell('J2')).toHaveText('20');
  await expect(cell('J6')).toHaveText('100');
  await expect(cell('J12')).toHaveText('120');
  await cell('B2').dblclick();
  await page.getByRole('textbox', { name: 'Edit B2', exact: true }).fill('50');
  await page.keyboard.press('Enter');
  await expect(cell('J2')).toHaveText('100');
  await expect(cell('J12')).toHaveText('168');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Calculations', exact: true }).click();
  await expect(cell('J12')).toHaveText('168');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
  const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
  const cached = (ref: string) =>
    xml
      .match(new RegExp(`<c\\b[^>]*r="${ref}"[^>]*>([\\s\\S]*?)</c>`))?.[1]
      .match(/<v>(.*?)<\/v>/)?.[1];
  expect(cached('J12')).toBe('168');
  expect(cached('L2')).toBe('336');
  await page.screenshot({ path: 'test-results/excel-table-calculation.png' });
});
