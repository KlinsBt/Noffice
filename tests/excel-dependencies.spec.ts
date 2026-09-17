import { test, expect } from '@playwright/test';
import ExcelJS from 'exceljs';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

test('Excel recalculates a transitive cross-sheet chain after edit, undo, redo and reload', async ({
  page,
}) => {
  const root = '.local/excel-dependencies';
  await fs.mkdir(root, { recursive: true });
  const book = new ExcelJS.Workbook();
  const data = book.addWorksheet('Data'),
    summary = book.addWorksheet('Summary');
  data.getCell('A1').value = 2;
  data.getCell('B1').value = { formula: 'A1*3', result: 6 };
  summary.getCell('A1').value = { formula: 'Data!B1+1', result: 7 };
  summary.getCell('B1').value = { formula: '2+2', result: 4 };
  const source = Buffer.from(await book.xlsx.writeBuffer());
  await fs.writeFile(`${root}/source.xlsx`, source);
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  await page.goto('/');
  const upload = async (buffer: Buffer, name: string) => {
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({
        name: `${name}.xlsx`,
        buffer,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  await upload(source, 'Dependencies');
  const cell = (name: string) => page.getByRole('gridcell', { name, exact: true });
  const tab = (name: string) => page.getByRole('button', { name, exact: true });
  await tab('Summary').click();
  await expect(cell('A1')).toHaveText('7');
  await expect(cell('B1')).toHaveText('4');
  await tab('Data').click();
  await cell('A1').dblclick();
  await page.getByRole('textbox', { name: 'Edit A1', exact: true }).fill('5');
  await page.keyboard.press('Enter');
  await expect(cell('B1')).toHaveText('15');
  await tab('Summary').click();
  await expect(cell('A1')).toHaveText('16');
  await expect(cell('B1')).toHaveText('4');
  await tab('Undo').click();
  await expect(cell('A1')).toHaveText('7');
  await tab('Redo').click();
  await expect(cell('A1')).toHaveText('16');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await tab('Recent files').click();
  await tab('Dependencies').click();
  await tab('Summary').click();
  await expect(cell('A1')).toHaveText('16');
  await tab('Export').click();
  const pending = page.waitForEvent('download');
  await tab('XLSX file Editable in Microsoft Excel').click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.writeFile(`${root}/browser.xlsx`, output);
  await upload(output, 'Dependencies reimport');
  await tab('Summary').click();
  await expect(cell('A1')).toHaveText('16');
  await expect(cell('B1')).toHaveText('4');
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceHash: hash(source),
        exportHash: hash(output),
        expected: { data: [5, 15], summary: [16, 4] },
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ),
  );
});
