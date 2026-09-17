import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { tableRenameFixture } from './fixtures/xlsx-table-rename';

test('table and column renaming repairs dependencies through undo, reload and consecutive XLSX exports', async ({
  page,
}) => {
  const root = '.local/xlsx-table-rename';
  await fs.mkdir(root, { recursive: true });
  const input = Buffer.from(await tableRenameFixture());
  await fs.writeFile(`${root}/source.xlsx`, input);
  await page.goto('/');
  await page
    .locator('input[type=file]')
    .first()
    .setInputFiles({
      name: 'Rename tables.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: input,
    });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const command = (name: string) => page.getByRole('button', { name, exact: true }).click();
  await expect(cell('E1')).toHaveText('10');
  await cell('B2').click();
  await page.keyboard.type('4');
  await page.keyboard.press('Enter');
  await expect(cell('E1')).toHaveText('12');
  await cell('B2').click();
  await command('Table design');
  await page.getByLabel('Table name', { exact: true }).fill('Orders');
  await page.getByLabel('Column name', { exact: true }).fill('Price');
  await command('Apply');
  await expect(page.getByRole('alert')).toHaveText('Another column already uses this name.');
  await page.getByLabel('Column name', { exact: true }).fill('Units sold');
  await command('Apply');
  await expect(page.getByRole('dialog')).not.toBeVisible();
  await expect(cell('B1')).toHaveText('Units sold');
  await expect(cell('E1')).toHaveText('12');
  await expect(cell('C2')).toHaveText('40');
  await expect(cell('H3')).toHaveText('Units sold');
  await expect(cell('H5')).toHaveText('Sales[Quantity]');
  await command('Undo');
  await expect(cell('B1')).toHaveText('Quantity');
  await expect(cell('B2')).toHaveText('4');
  await command('Redo');
  await expect(cell('B1')).toHaveText('Units sold');
  const download = async (stage: string) => {
    await command('Export');
    const pending = page.waitForEvent('download');
    await command('XLSX file Editable in Microsoft Excel');
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${stage}.xlsx`, bytes);
    return bytes;
  };
  await download('renamed');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await command('Recent files');
  await command('Rename tables');
  await expect(cell('B1')).toHaveText('Units sold');
  await expect(cell('E1')).toHaveText('12');
  await cell('B2').click();
  await command('Table design');
  await expect(page.getByLabel('Table name', { exact: true })).toHaveValue('Orders');
  await page.getByLabel('Table name', { exact: true }).fill('Revenue');
  await page.getByLabel('Column name', { exact: true }).fill('Net, total');
  await command('Apply');
  await expect(cell('B1')).toHaveText('Net, total');
  await expect(cell('E1')).toHaveText('12');
  await download('special');
  await cell('B2').click();
  await command('Table design');
  await page.getByLabel('Column name', { exact: true }).fill("#[Q]@'");
  await command('Apply');
  await expect(cell('B1')).toHaveText("#[Q]@'");
  await expect(cell('E1')).toHaveText('12');
  const bytes = await download('escaped'),
    zip = await JSZip.loadAsync(bytes),
    original = await JSZip.loadAsync(input);
  for (const path of [
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/comments1.xml',
    'xl/worksheets/sheet2.xml',
    'custom/preservation.xml',
  ])
    expect(await zip.file(path)!.async('string')).toBe(await original.file(path)!.async('string'));
  await cell('B2').click();
  await command('Table design');
  await page.screenshot({ path: 'test-results/excel-table-rename.png' });
  await command('Cancel');
  await page
    .locator('input[type=file]')
    .first()
    .setInputFiles({
      name: 'Renamed export.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: bytes,
    });
  await expect(cell('B1')).toHaveText("#[Q]@'");
  await expect(cell('E1')).toHaveText('12');
  await expect(cell('H3')).toHaveText("#[Q]@'");
});
