import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { tableResizeFixture } from './fixtures/xlsx-table-resize';

test('calculated tables grow, shrink and regrow through undo, reload and preserved exports', async ({
  page,
}) => {
  const root = '.local/xlsx-table-resize';
  await fs.mkdir(root, { recursive: true });
  const input = Buffer.from(await tableResizeFixture());
  await fs.writeFile(`${root}/source.xlsx`, input);
  await page.goto('/');
  const open = (name: string, buffer: Buffer) =>
    page.locator('input[type=file]').first().setInputFiles({
      name,
      buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  await open('Resize.xlsx', input);
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const command = (name: string) => page.getByRole('button', { name, exact: true }).click();
  const resize = async (ref: string) => {
    await cell('B2').click();
    await command('Table design');
    await page.getByLabel('Table range', { exact: true }).fill(ref);
    await command('Apply');
    await expect(page.getByLabel('Table range', { exact: true })).toHaveCount(0);
  };
  const download = async (stage: string) => {
    await command('Export');
    const pending = page.waitForEvent('download');
    await command('XLSX file Editable in Microsoft Excel');
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${stage}.xlsx`, bytes);
    return bytes;
  };
  await cell('C3').click();
  await page.keyboard.type('99');
  await page.keyboard.press('Enter');
  await resize('A1:C6');
  await expect(cell('C5')).toHaveText('40');
  await expect(cell('C6')).toHaveText('110');
  await expect(cell('H4')).toHaveText('347');
  await download('grown');
  await command('Undo');
  await expect(cell('C6')).toHaveText('');
  await expect(cell('H4')).toHaveText('179');
  await command('Redo');
  await expect(cell('C6')).toHaveText('110');
  await resize('A1:C3');
  await expect(cell('C4')).toHaveText('#VALUE!');
  await expect(cell('C6')).toHaveText('#VALUE!');
  await expect(cell('H4')).toHaveText('124');
  await expect(cell('C5')).toHaveText('40');
  await download('shrunk');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await command('Recent files');
  await command('Resize');
  await expect(cell('C6')).toHaveText('#VALUE!');
  await resize('A1:C6');
  await expect(cell('C6')).toHaveText('110');
  const bytes = await download('regrown'),
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
  await open('Reimport resized.xlsx', bytes);
  await expect(cell('H4')).toHaveText('347');
  await cell('C2').click();
  await cell('C6').click({ modifiers: ['Shift'] });
  await page.keyboard.press('Delete');
  await expect(cell('C2')).toHaveText('');
  await expect(cell('H4')).toHaveText('28');
  const cleared = await JSZip.loadAsync(await download('cleared'));
  expect(await cleared.file('xl/tables/table1.xml')!.async('string')).not.toContain(
    'calculatedColumnFormula',
  );
  await command('Undo');
  await expect(cell('C6')).toHaveText('110');
  await command('Redo');
  await expect(cell('C6')).toHaveText('');
  await resize('A1:C8');
  await expect(cell('C7')).toHaveText('');
  await expect(cell('C8')).toHaveText('');
  await download('empty-grown');
});

test('new calculated tables resize with relative formulas and keep excluded worksheet cells', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Excel', exact: true }).click();
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const command = (name: string) => page.getByRole('button', { name, exact: true }).click();
  const enter = async (ref: string, value: string) => {
    await cell(ref).click();
    await page.keyboard.type(value);
    await page.keyboard.press('Enter');
  };
  for (const [ref, value] of [
    ['A1', 'Quantity'],
    ['B1', 'Total'],
    ['A2', '2'],
    ['A3', '3'],
  ])
    await enter(ref, value);
  await cell('A1').click();
  await cell('B3').click({ modifiers: ['Shift'] });
  await command('Format as table');
  await page.getByLabel('Table name', { exact: true }).fill('Sales');
  await command('Create table');
  await enter('B2', '=A2*10+$A$2');
  const resize = async (ref: string) => {
    await cell('A2').click();
    await command('Table design');
    await page.getByLabel('Table range', { exact: true }).fill(ref);
    await command('Apply');
  };
  const download = async (stage: string) => {
    await command('Export');
    const pending = page.waitForEvent('download');
    await command('XLSX file Editable in Microsoft Excel');
    await fs.mkdir('.local/xlsx-table-resize', { recursive: true });
    await fs.writeFile(
      `.local/xlsx-table-resize/${stage}.xlsx`,
      await fs.readFile((await (await pending).path())!),
    );
  };
  await resize('A1:B5');
  await enter('A4', '7');
  await enter('A5', '11');
  await expect(cell('B4')).toHaveText('72');
  await expect(cell('B5')).toHaveText('112');
  await download('new-grown');
  await resize('A1:B2');
  await expect(cell('B3')).toHaveText('32');
  await expect(cell('B5')).toHaveText('112');
  await download('new-shrunk');
});
