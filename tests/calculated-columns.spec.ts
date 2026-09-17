import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { calculatedColumnsFixture } from './fixtures/xlsx-calculated-columns';

test('calculated columns fill hidden rows, update atomically and preserve exceptions through rename, reload and export', async ({
  page,
}) => {
  const root = '.local/xlsx-calculated-columns';
  await fs.mkdir(root, { recursive: true });
  const input = Buffer.from(await calculatedColumnsFixture());
  await fs.writeFile(`${root}/source.xlsx`, input);
  await page.goto('/');
  const open = (name: string, buffer: Buffer) =>
    page.locator('input[type=file]').first().setInputFiles({
      name,
      buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  await open('Calculated.xlsx', input);
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const command = (name: string) => page.getByRole('button', { name, exact: true }).click();
  const enter = async (ref: string, value: string) => {
    await cell(ref).click();
    await page.keyboard.type(value);
    await page.keyboard.press('Enter');
  };
  const totals = async (a: number, b: number, sum: number) => {
    await expect(cell('C2')).toHaveText(String(a));
    await expect(cell('C3')).toHaveText(String(b));
    await expect(cell('E1')).toHaveText(String(sum));
    await expect(cell('E2')).toHaveText(String(sum * 2));
    await expect(cell('C5')).toHaveText('40');
    await expect(cell('C4')).toHaveCount(0);
  };
  const download = async (stage: string) => {
    await command('Export');
    const pending = page.waitForEvent('download');
    await command('XLSX file Editable in Microsoft Excel');
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${stage}.xlsx`, bytes);
    return bytes;
  };
  await enter('C3', '=Sales[[#This Row],[Quantity]]*10');
  await totals(20, 30, 100);
  await download('structured');
  await command('Undo');
  await expect(cell('C2')).toHaveText('');
  await expect(cell('E1')).toHaveText('0');
  await command('Redo');
  await totals(20, 30, 100);
  await enter('C2', '=Sales[[#This Row],[Quantity]]*20');
  await totals(40, 60, 200);
  await download('updated');
  await enter('B1', 'Units');
  await expect(cell('B1')).toHaveText('Units');
  await totals(40, 60, 200);
  await download('renamed');
  await enter('C3', '99');
  await totals(40, 99, 239);
  await download('exception');
  await enter('C2', '=Sales[[#This Row],[Units]]*30');
  await totals(60, 99, 259);
  await download('isolated');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await command('Recent files');
  await command('Calculated');
  await totals(60, 99, 259);
  await cell('C2').click();
  await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData('text/plain', '\n\n\n');
    document.activeElement!.dispatchEvent(
      new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
    );
  });
  await expect(cell('E1')).toHaveText('0');
  await enter('C3', '=B3*10+$B$2');
  await totals(22, 32, 106);
  const a1 = await download('a1');
  await open('Calculated reimport.xlsx', a1);
  await totals(22, 32, 106);
  await cell('C2').click();
  await page.getByLabel('Formula bar', { exact: true }).fill('=B2*20+$B$2');
  await page.getByLabel('Formula bar', { exact: true }).press('Enter');
  await totals(42, 62, 206);
  const output = await download('a1updated'),
    zip = await JSZip.loadAsync(output),
    original = await JSZip.loadAsync(input);
  expect(await zip.file('xl/tables/table1.xml')!.async('string')).toContain(
    '<calculatedColumnFormula>B2*20+$B$2</calculatedColumnFormula>',
  );
  for (const path of [
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/comments1.xml',
    'xl/worksheets/sheet2.xml',
    'custom/preservation.xml',
  ])
    expect(await zip.file(path)!.async('string')).toBe(await original.file(path)!.async('string'));
  await command('Undo');
  await totals(22, 32, 106);
  await command('Redo');
  await totals(42, 62, 206);
});

test('a new table calculates from direct entry and retains its master formula after XLSX reimport', async ({
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
    ['A4', '5'],
  ])
    await enter(ref, value);
  await cell('A1').click();
  await cell('B4').click({ modifiers: ['Shift'] });
  await command('Format as table');
  await page.getByLabel('Table name', { exact: true }).fill('Sales');
  await command('Create table');
  await enter('B3', '=A3*10');
  await expect(cell('B2')).toHaveText('20');
  await expect(cell('B4')).toHaveText('50');
  await enter('B4', '99');
  await enter('B3', '=A3*20');
  await expect(cell('B2')).toHaveText('20');
  await expect(cell('B3')).toHaveText('60');
  await expect(cell('B4')).toHaveText('99');
  await command('Export');
  const pending = page.waitForEvent('download');
  await command('XLSX file Editable in Microsoft Excel');
  const bytes = await fs.readFile((await (await pending).path())!);
  await fs.mkdir('.local/xlsx-calculated-columns', { recursive: true });
  await fs.writeFile('.local/xlsx-calculated-columns/new.xlsx', bytes);
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'New calculated.xlsx',
    buffer: bytes,
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  await expect(cell('B2')).toHaveText('20');
  await expect(cell('B3')).toHaveText('60');
  await expect(cell('B4')).toHaveText('99');
});
