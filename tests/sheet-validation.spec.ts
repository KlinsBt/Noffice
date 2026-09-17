import { expect, test } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { validationFixture } from './fixtures/xlsx-validation';
test('Excel edits validation with blank-range enforcement, undo, persistence and retained export', async ({
  page,
}) => {
  const input = Buffer.from(await validationFixture());
  await fs.mkdir('.local/xlsx-validation', { recursive: true });
  await fs.writeFile('.local/xlsx-validation/validation-source.xlsx', input);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Validation workbook.xlsx',
    mimeType: 'application/octet-stream',
    buffer: input,
  });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const open = () => page.getByRole('button', { name: 'Data validation', exact: true }).click();
  const apply = () => page.getByRole('button', { name: 'Apply validation', exact: true }).click();
  await cell('A3').click();
  await page.keyboard.type('2');
  await page.keyboard.press('Enter');
  await expect(page.getByText('A3: Use the allowed range', { exact: true })).toBeVisible();
  await expect(cell('A3')).toHaveText('');
  await cell('A2').click();
  await open();
  await page.getByLabel('Validation type').selectOption('decimal');
  await page.getByLabel('Validation first value').fill('0');
  await page.getByLabel('Validation second value').fill('1');
  await page.getByLabel('Validation input title').fill('Fraction');
  await page.getByLabel('Validation input message').fill('Enter a fraction');
  await page.getByLabel('Validation error title').fill('Invalid fraction');
  await page.getByLabel('Validation error message').fill('Between zero and one');
  await apply();
  // Adding a rule does not retroactively discard an existing invalid value.
  await expect(cell('A2')).toHaveText('5');
  await page.getByRole('button', { name: 'Bold', exact: true }).click();
  await expect(cell('A2')).toHaveText('5');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await open();
  await expect(page.getByLabel('Validation type')).toHaveValue('whole');
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await cell('A2').click();
  await page.keyboard.type('0.5');
  await page.keyboard.press('Enter');
  await expect(cell('A2')).toHaveText('0.5');
  await cell('A2').click();
  await page.keyboard.type('2');
  await page.keyboard.press('Enter');
  await expect(page.getByText('A2: Between zero and one', { exact: true })).toBeVisible();
  await expect(cell('A2')).toHaveText('0.5');
  await cell('D2').click();
  await expect(page.getByLabel('Cell choices')).toBeVisible();
  await open();
  await page.getByLabel('Validation first value').fill('"Blue,Orange"');
  await apply();
  await page.getByLabel('Cell choices').selectOption('Blue');
  await expect(cell('D2')).toHaveText('Blue');
  await cell('D4').click();
  await open();
  await page.getByRole('button', { name: 'Clear validation', exact: true }).click();
  await expect(page.getByLabel('Cell choices')).toHaveCount(0);
  await cell('D3').click();
  await expect(page.getByLabel('Cell choices').locator('option')).toHaveText([
    'Choose a value',
    'Red',
    'Green',
  ]);
  await page.getByRole('button', { name: 'Protected', exact: true }).click();
  await open();
  await apply();
  await expect(page.getByRole('alert')).toContainText('protected');
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  await page.getByRole('button', { name: 'Entry', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Validation workbook', exact: true }).click();
  await expect(cell('A2')).toHaveText('0.5');
  await expect(cell('E1')).toHaveText('1');
  await cell('D4').click();
  await expect(page.getByLabel('Cell choices')).toHaveCount(0);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.writeFile('.local/xlsx-validation/validation-edited.xlsx', output);
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && !['xl/worksheets/sheet1.xml', 'xl/workbook.xml'].includes(path))
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Validation result.xlsx',
    mimeType: 'application/octet-stream',
    buffer: output,
  });
  await cell('A2').click();
  await open();
  await expect(page.getByLabel('Validation type')).toHaveValue('decimal');
  await expect(page.getByLabel('Validation input message')).toHaveValue('Enter a fraction');
  await page.screenshot({ path: 'test-results/excel-validation.png' });
});
