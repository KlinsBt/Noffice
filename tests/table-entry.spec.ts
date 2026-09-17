import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { tableEntryFixture } from './fixtures/xlsx-table-entry';
test('typing below a table expands it atomically and preserves formula-entry and occupied-row behavior', async ({
  page,
}) => {
  const root = '.local/xlsx-table-entry';
  await fs.mkdir(root, { recursive: true });
  const input = Buffer.from(await tableEntryFixture());
  await fs.writeFile(`${root}/source.xlsx`, input);
  await page.goto('/');
  const open = (name: string, buffer: Buffer) =>
    page.locator('input[type=file]').first().setInputFiles({
      name,
      buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  await open('Table entry.xlsx', input);
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const command = (name: string) => page.getByRole('button', { name, exact: true }).click();
  const enter = async (ref: string, value: string) => {
    await cell(ref).click();
    await page.keyboard.type(value);
    await page.keyboard.press('Enter');
  };
  const range = async (ref: string) => {
    await cell('A2').click();
    await command('Table design');
    await expect(page.getByLabel('Table range', { exact: true })).toHaveValue(ref);
    await command('Cancel');
  };
  const download = async (stage: string) => {
    await command('Export');
    const pending = page.waitForEvent('download');
    await command('XLSX file Editable in Microsoft Excel');
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${stage}.xlsx`, bytes);
    return bytes;
  };
  await enter('A5', 'East');
  await range('A1:C5');
  await expect(cell('C5')).toHaveText('0');
  await download('first');
  await command('Undo');
  await range('A1:C4');
  await expect(cell('A5')).toHaveText('');
  await expect(cell('C5')).toHaveText('');
  await command('Redo');
  await range('A1:C5');
  await expect(cell('A5')).toHaveText('East');
  await enter('B5', '7');
  await expect(cell('C5')).toHaveText('70');
  await expect(cell('H4')).toHaveText('187');
  await download('quantity');
  await enter('B6', '11');
  await range('A1:C6');
  await expect(cell('C6')).toHaveText('110');
  await expect(cell('H4')).toHaveText('308');
  await download('next');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await command('Recent files');
  await command('Table entry');
  await range('A1:C6');
  await cell('B7').click();
  await page.getByLabel('Formula bar', { exact: true }).fill('=2+3');
  await page.getByLabel('Formula bar', { exact: true }).press('Enter');
  await range('A1:C7');
  await expect(cell('C7')).toHaveText('');
  await expect(cell('H4')).toHaveText('308');
  await expect(cell('A7')).toHaveText('Total');
  await download('formula');
  await command('Undo');
  await range('A1:C6');
  await expect(cell('A7')).toHaveText('');
  await command('Redo');
  await range('A1:C7');
  await expect(cell('A7')).toHaveText('Total');
  await enter('C7', '99');
  await expect(cell('H4')).toHaveText('308');
  await download('exception');
  await cell('C8').click();
  await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData('text/plain', '55');
    document.activeElement!.dispatchEvent(
      new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
    );
  });
  await enter('B8', '13');
  await range('A1:C7');
  await expect(cell('C8')).toHaveText('55');
  await expect(cell('H4')).toHaveText('308');
  const output = await download('neighbor'),
    zip = await JSZip.loadAsync(output),
    original = await JSZip.loadAsync(input);
  for (const path of [
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/comments1.xml',
    'xl/worksheets/sheet2.xml',
    'custom/preservation.xml',
  ])
    expect(await zip.file(path)!.async('string')).toBe(await original.file(path)!.async('string'));
  const tableXml = await zip.file('xl/tables/table1.xml')!.async('string');
  expect(tableXml).toContain('totalsRowCount="1"');
  expect(tableXml).toContain('<totalsRowFormula>2+3</totalsRowFormula>');
  expect(tableXml).toContain('totalsRowLabel="99"');
  await open('Reimport.xlsx', output);
  await range('A1:C7');
  await expect(cell('H4')).toHaveText('308');
});
