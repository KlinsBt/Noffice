import { expect, test } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { structureFixture } from './fixtures/xlsx-structure';
import { validationFixture } from './fixtures/xlsx-validation';

test('new Excel workbooks insert multiple selected rows and still allow adding sheets', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Excel', exact: true }).click();
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  for (const [ref, value] of [
    ['A1', '2'],
    ['A2', '3'],
    ['B1', '=SUM(A1:A2)'],
  ]) {
    await cell(ref).click();
    await page.keyboard.type(value);
    await page.keyboard.press('Enter');
  }
  await cell('A2').click();
  await page.keyboard.press('Shift+ArrowDown');
  await page.getByLabel('Insert or delete rows and columns').selectOption('insert-row');
  await expect(cell('A4')).toHaveText('3');
  await expect(cell('B1')).toHaveText('5');
  await page.getByRole('button', { name: 'Add sheet', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Sheet 2', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Sheet 1', exact: true }).click();
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  const zip = await JSZip.loadAsync(await fs.readFile((await (await pending).path())!));
  expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toContain(
    '<f>SUM(A1:A4)</f>',
  );
  expect(await zip.file('xl/workbook.xml')!.async('string')).toContain('name="Sheet 2"');
});

test('Excel inserts and deletes whole rows and columns with formula repair, undo, saved checkpoints and retained export', async ({
  page,
}) => {
  const input = Buffer.from(await structureFixture());
  await fs.mkdir('.local/xlsx-structure', { recursive: true });
  await fs.writeFile('.local/xlsx-structure/source.xlsx', input);
  await page.goto('/');
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles({ name: 'Structure.xlsx', mimeType: 'application/octet-stream', buffer: input });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const command = (name: string) =>
    /^(Insert|Delete) (rows|columns)$/.test(name)
      ? page.getByLabel('Insert or delete rows and columns').selectOption({ label: name })
      : page.getByRole('button', { name, exact: true }).click();
  const download = async (name: string) => {
    await command('Export');
    const pending = page.waitForEvent('download');
    await command('XLSX file Editable in Microsoft Excel');
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`.local/xlsx-structure/${name}.xlsx`, bytes);
    return bytes;
  };
  await cell('A3').click();
  await command('Insert rows');
  await expect(cell('A4')).toHaveText('20');
  await cell('B2').click();
  await expect(page.getByLabel('Formula bar')).toHaveValue('=A2+$A$5');
  await download('insert-row');
  await command('Undo');
  await expect(cell('A3')).toHaveText('20');
  await command('Redo');
  await expect(cell('A4')).toHaveText('20');
  await cell('B2').click();
  await command('Insert columns');
  await expect(cell('C2')).toHaveText('40');
  await download('insert-column');
  await cell('A4').click();
  await command('Delete rows');
  await expect(cell('A4')).toHaveText('30');
  await download('delete-row');
  await cell('D1').click();
  await command('Delete columns');
  await expect(cell('D1')).toHaveText('#REF!');
  await cell('A2').click();
  await page.keyboard.type('12');
  await page.keyboard.press('Enter');
  await expect(cell('C2')).toHaveText('42');
  await command("O'Brien");
  await expect(cell('A1')).toHaveText('#REF!');
  await expect(cell('A2')).toHaveText('42');
  await expect(cell('A3')).toHaveText('132');
  await command('Data');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await command('Recent files');
  await command('Structure');
  await expect(cell('C2')).toHaveText('42');
  const output = await download('edited');
  await cell('A1').click();
  await command('Insert rows');
  await expect(cell('A3')).toHaveText('12.00');
  await download('insert-first-row');
  await cell('A1').click();
  await command('Insert columns');
  await expect(cell('B3')).toHaveText('12.00');
  await download('insert-first-column');
  await cell('A1').click();
  await command('Delete rows');
  await expect(cell('B2')).toHaveText('12.00');
  await download('delete-first-row');
  await cell('A1').click();
  await command('Delete columns');
  await expect(cell('A2')).toHaveText('12.00');
  await download('delete-first-column');
  const source = await JSZip.loadAsync(input),
    exported = await JSZip.loadAsync(output);
  for (const path of [
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/worksheets/sheet3.xml',
    'custom/preservation.xml',
  ])
    expect(await exported.file(path)!.async('uint8array'), path).toEqual(
      await source.file(path)!.async('uint8array'),
    );
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Reopened structure.xlsx',
    mimeType: 'application/octet-stream',
    buffer: output,
  });
  await expect(cell('C2')).toHaveText('42');
  await cell('C2').click();
  await expect(page.getByLabel('Formula bar')).toHaveValue('=A2+$A$4');
  await page.screenshot({ path: 'test-results/excel-structure.png' });
});

test('Excel guards unsupported worksheet anchors and protection without changing the workbook', async ({
  page,
}) => {
  const input = Buffer.from(await validationFixture());
  await page.goto('/');
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles({ name: 'Guarded.xlsx', mimeType: 'application/octet-stream', buffer: input });
  await page.getByRole('gridcell', { name: 'A2', exact: true }).click();
  await page.getByLabel('Insert or delete rows and columns').selectOption('insert-row');
  await expect(
    page.getByText(
      'This worksheet contains protection, drawings, notes or other anchors that cannot yet be repaired by row and column operations.',
      { exact: true },
    ),
  ).toBeVisible();
  await expect(page.getByRole('gridcell', { name: 'A2', exact: true })).toHaveText('5');
  await page.getByRole('button', { name: 'Protected', exact: true }).click();
  await expect(page.getByLabel('Insert or delete rows and columns')).toBeDisabled();
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  expect(await fs.readFile((await (await pending).path())!)).toEqual(input);
});
