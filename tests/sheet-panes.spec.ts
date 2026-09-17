import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import fs from 'node:fs/promises';

test('Excel freezes headings while virtualizing the body and preserves pane edits through undo, reload and XLSX exports', async ({
  page,
}) => {
  const book = new ExcelJS.Workbook(),
    sheet = book.addWorksheet('Pane test');
  sheet.views = [{ state: 'normal', zoomScale: 85 }];
  for (let r = 1; r <= 2000; r++)
    for (let c = 1; c <= 12; c++) sheet.getCell(r, c).value = `Row ${r} column ${c}`;
  sheet.getRow(1).height = 30;
  sheet.getRow(2).hidden = true;
  sheet.getRow(3).height = 24;
  sheet.getColumn(1).width = 22;
  sheet.getColumn(2).width = 18;
  sheet.getCell('D4').value = { formula: '1+2', result: 3 };
  sheet.getCell('D4').note = 'Keep this note';
  const original = Buffer.from(await book.xlsx.writeBuffer());
  await page.setViewportSize({ width: 1440, height: 1000 });
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Pane test.xlsx',
    mimeType: 'application/octet-stream',
    buffer: original,
  });
  await page.getByRole('gridcell', { name: 'C4', exact: true }).click();
  const panes = page.getByRole('combobox', { name: 'Freeze panes', exact: true });
  await panes.selectOption('selection');
  await expect(page.getByText('3 rows · 2 columns frozen', { exact: true })).toBeVisible();
  const first = page.getByRole('gridcell', { name: 'A1', exact: true }),
    second = page.getByRole('gridcell', { name: 'A3', exact: true });
  const beforeFirst = (await first.boundingBox())!,
    beforeSecond = (await second.boundingBox())!;
  const grid = page.getByRole('region', { name: 'Spreadsheet grid', exact: true });
  await grid.evaluate((el) => {
    el.scrollTop = 9000;
    el.scrollLeft = 400;
  });
  await expect.poll(() => page.locator('tbody tr[aria-rowindex]').count()).toBeLessThan(65);
  await expect(page.getByRole('gridcell', { name: 'A4', exact: true })).toHaveCount(0);
  const afterFirst = (await first.boundingBox())!,
    afterSecond = (await second.boundingBox())!;
  expect(afterFirst.x).toBeCloseTo(beforeFirst.x, 0);
  expect(afterFirst.y).toBeCloseTo(beforeFirst.y, 0);
  expect(afterSecond.x).toBeCloseTo(beforeSecond.x, 0);
  expect(afterSecond.y).toBeCloseTo(beforeSecond.y, 0);
  const header = page.getByRole('columnheader', { name: 'A', exact: true });
  expect((await header.boundingBox())!.x).toBeCloseTo(afterFirst.x, 0);
  await page.screenshot({ path: 'test-results/excel-frozen-panes.png' });
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(first).toHaveCount(0);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(first).toHaveCSS('position', 'sticky');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Pane test', exact: true }).click();
  await expect(page.getByText('3 rows · 2 columns frozen', { exact: true })).toBeVisible();
  async function download(name: string) {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.mkdir('.local/xlsx-validation', { recursive: true });
    await fs.writeFile(`.local/xlsx-validation/${name}.xlsx`, bytes);
    return bytes;
  }
  const frozen = await download('frozen-panes'),
    before = await JSZip.loadAsync(original),
    after = await JSZip.loadAsync(frozen);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'xl/worksheets/sheet1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const checked = new ExcelJS.Workbook();
  await checked.xlsx.load(frozen as any);
  expect(checked.worksheets[0].views[0]).toMatchObject({
    state: 'frozen',
    xSplit: 2,
    ySplit: 3,
    topLeftCell: 'C4',
    zoomScale: 85,
  });
  expect(checked.worksheets[0].getCell('D4').value).toEqual({ formula: '1+2', result: 3 });
  await panes.selectOption('row');
  await expect(page.getByText('1 rows · 0 columns frozen', { exact: true })).toBeVisible();
  await panes.selectOption('column');
  await expect(page.getByText('0 rows · 1 columns frozen', { exact: true })).toBeVisible();
  await panes.selectOption('none');
  await expect(first).not.toHaveCSS('position', 'sticky');
  const unfrozen = await download('unfrozen-panes');
  const normal = new ExcelJS.Workbook();
  await normal.xlsx.load(unfrozen as any);
  expect(normal.worksheets[0].views[0].state).toBe('normal');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Frozen result.xlsx',
    mimeType: 'application/octet-stream',
    buffer: frozen,
  });
  await expect(page.getByText('3 rows · 2 columns frozen', { exact: true })).toBeVisible();
});
