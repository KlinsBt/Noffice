import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { tableRenameFixture } from './fixtures/xlsx-table-rename';
test('Unicode and backslash table names calculate, persist and export through consecutive renames', async ({
  page,
}) => {
  const root = '.local/xlsx-table-names';
  await fs.mkdir(root, { recursive: true });
  const input = Buffer.from(await tableRenameFixture());
  await fs.writeFile(`${root}/source.xlsx`, input);
  const names = ['Ümsatz', '日本', '\\Sales', 'Sales\\West'];
  await fs.writeFile(`${root}/names.json`, JSON.stringify(names));
  await page.goto('/');
  await page
    .locator('input[type=file]')
    .first()
    .setInputFiles({
      name: 'Table names.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: input,
    });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const command = (name: string) => page.getByRole('button', { name, exact: true }).click();
  let last = Buffer.alloc(0);
  for (const [i, name] of names.entries()) {
    await cell('B2').click();
    await command('Table design');
    await page.getByLabel('Table name', { exact: true }).fill(name);
    await command('Apply');
    await expect(page.getByRole('dialog')).not.toBeVisible();
    await expect(cell('E1')).toHaveText('10');
    await expect(cell('C2')).toHaveText('20');
    await command('Export');
    const pending = page.waitForEvent('download');
    await command('XLSX file Editable in Microsoft Excel');
    last = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/stage-${i}.xlsx`, last);
  }
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await command('Recent files');
  await command('Table names');
  await expect(cell('E1')).toHaveText('10');
  await cell('B2').click();
  await command('Table design');
  await expect(page.getByLabel('Table name', { exact: true })).toHaveValue('Sales\\West');
  await command('Cancel');
  await page
    .locator('input[type=file]')
    .first()
    .setInputFiles({
      name: 'Names export.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: last,
    });
  await expect(cell('E1')).toHaveText('10');
});
