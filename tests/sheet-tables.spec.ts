import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { tableFixture } from './fixtures/xlsx-tables';
import type { OfficeFile } from '../src/model';

test('older saved workbook tables recover metadata on open without a storage revision or lost edits', async ({
  page,
}) => {
  const input = Buffer.from(await tableFixture());
  await page.goto('/');
  await page
    .locator('input[type=file]')
    .first()
    .setInputFiles({
      name: 'Legacy tables.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: input,
    });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  await expect(cell('A1')).toHaveText('Region');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.evaluate(
    () =>
      new Promise<void>((resolve, reject) => {
        const request = indexedDB.open('noffice-workspace');
        request.onsuccess = () => {
          const db = request.result,
            tx = db.transaction('files', 'readwrite'),
            store = tx.objectStore('files'),
            all = store.getAll();
          all.onsuccess = () => {
            const file = all.result.find(
              (f: OfficeFile) => f.name === 'Legacy tables',
            ) as OfficeFile;
            if (file.content.kind !== 'excel') throw Error('Expected workbook');
            file.content.sheets[0].cells.B2 = { value: '15' };
            for (const s of file.content.sheets) {
              delete s.tableTheme;
              for (const t of s.tables || []) {
                delete t.sourcePath;
                delete t.style;
                delete t.resizeBlocked;
              }
            }
            store.put(file);
          };
          tx.oncomplete = () => {
            db.close();
            resolve();
          };
          tx.onerror = () => reject(tx.error);
        };
        request.onerror = () => reject(request.error);
      }),
  );
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Legacy tables', exact: true }).click();
  await expect(cell('A1')).toHaveCSS('background-color', 'rgb(79, 129, 189)');
  await expect(cell('E1')).toHaveText('23');
  await cell('A2').click();
  await page.getByRole('button', { name: 'Table design', exact: true }).click();
  await page.getByLabel('Table range', { exact: true }).fill('A1:C6');
  await page.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(cell('E1')).toHaveText('41');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Legacy tables', exact: true }).click();
  await expect(cell('E1')).toHaveText('41');
});

test('Excel table creation, resizing, styles, filters, undo and saved exports preserve structured calculations', async ({
  page,
}) => {
  const root = '.local/xlsx-tables';
  await fs.mkdir(root, { recursive: true });
  const input = Buffer.from(await tableFixture());
  await fs.writeFile(`${root}/source.xlsx`, input);
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'Tables.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: input,
  });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const command = (name: string) => page.getByRole('button', { name, exact: true }).click();
  const download = async (stage: string) => {
    await command('Export');
    const pending = page.waitForEvent('download');
    await command('XLSX file Editable in Microsoft Excel');
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${stage}.xlsx`, bytes);
    return bytes;
  };
  await expect(cell('E1')).toHaveText('10');
  await expect(cell('A1')).toHaveCSS('background-color', 'rgb(79, 129, 189)');
  await cell('A2').click();
  await command('Table design');
  await page.getByLabel('Table range', { exact: true }).fill('A1:C6');
  await command('Apply');
  await expect(cell('E1')).toHaveText('28');
  await expect(cell('E2')).toHaveText('56');
  await download('resized');
  await command('Undo');
  await expect(cell('E1')).toHaveText('10');
  await command('Redo');
  await expect(cell('E1')).toHaveText('28');
  const colors = [];
  for (let n = 2; n <= 7; n++) {
    await command('Table design');
    await page.getByLabel('Table style', { exact: true }).selectOption(`TableStyleMedium${n}`);
    await command('Apply');
    await download(`style-${n}`);
    for (const ref of ['A1', 'A2', 'A3'])
      colors.push({
        style: n,
        ref,
        ...(await cell(ref).evaluate((el) => ({
          fill: getComputedStyle(el).backgroundColor,
          color: getComputedStyle(el).color,
          bold: Number(getComputedStyle(el).fontWeight) >= 600,
        }))),
      });
  }
  for (const stage of ['bands', 'none']) {
    await command('Table design');
    if (stage === 'bands') {
      await page.getByLabel('Banded rows', { exact: true }).uncheck();
      await page.getByLabel('Banded columns', { exact: true }).check();
      await page.getByLabel('First column', { exact: true }).check();
      await page.getByLabel('Last column', { exact: true }).check();
    } else await page.getByLabel('Table style', { exact: true }).selectOption('');
    await command('Apply');
    await download(stage);
    for (const ref of stage === 'bands'
      ? ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3']
      : ['A1', 'A2', 'A3'])
      colors.push({
        style: stage,
        ref,
        ...(await cell(ref).evaluate((el) => ({
          fill: getComputedStyle(el).backgroundColor,
          color: getComputedStyle(el).color,
          bold: Number(getComputedStyle(el).fontWeight) >= 600,
        }))),
      });
  }
  await fs.writeFile(`${root}/browser-colors.json`, JSON.stringify(colors, null, 2));
  await command('Table design');
  await page.getByLabel('Table style', { exact: true }).selectOption('TableStyleMedium7');
  await page.getByLabel('Banded rows', { exact: true }).check();
  await page.getByLabel('Banded columns', { exact: true }).uncheck();
  await page.getByLabel('Last column', { exact: true }).uncheck();
  await page.getByLabel('Table range', { exact: true }).fill('A1:C4');
  await page.getByLabel('First column', { exact: true }).check();
  await command('Apply');
  await expect(cell('E1')).toHaveText('10');
  await expect(cell('B5')).toHaveText('7');
  await download('shrunk');
  await cell('F3').click();
  await cell('G6').click({ modifiers: ['Shift'] });
  await command('Format as table');
  await page.getByLabel('Table name', { exact: true }).fill('Supplies');
  await command('Create table');
  await cell('E4').click();
  await page.keyboard.type('=SUM(Supplies[Amount])');
  await page.keyboard.press('Enter');
  await expect(cell('E4')).toHaveText('40');
  await cell('A2').click();
  await command('Filter range');
  await page.getByRole('combobox', { name: 'Filter column' }).selectOption({ label: 'Region' });
  await page.getByRole('checkbox', { name: 'South', exact: true }).uncheck();
  await command('Apply filter');
  await expect(cell('A3')).toHaveCount(0);
  await command('Table design');
  await page.getByLabel('Table range', { exact: true }).fill('A1:C6');
  await command('Apply');
  await expect(page.getByRole('alert')).toHaveText(
    'Clear filter criteria before resizing the table.',
  );
  await command('Cancel');
  await command('Clear filters');
  await expect(cell('A3')).toHaveText('South');
  await cell('F4').click();
  await command('Table design');
  await page.getByLabel('Table style', { exact: true }).selectOption('TableStyleMedium4');
  await page.getByLabel('Last column', { exact: true }).check();
  await command('Apply');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await command('Recent files');
  await command('Tables');
  await expect(cell('E4')).toHaveText('40');
  const output = await download('created'),
    zip = await JSZip.loadAsync(output),
    original = await JSZip.loadAsync(input);
  for (const path of [
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/worksheets/sheet2.xml',
    'custom/preservation.xml',
    'xl/comments1.xml',
  ])
    expect(await zip.file(path)!.async('string')).toBe(await original.file(path)!.async('string'));
  await cell('A1').click();
  await page.keyboard.type('New header');
  await page.keyboard.press('Enter');
  await expect(cell('A1')).toHaveText('New header');
  await command('Undo');
  await expect(cell('A1')).toHaveText('Region');
  await page.setViewportSize({ width: 1440, height: 900 });
  await cell('F4').click();
  await command('Table design');
  await page.screenshot({ path: 'test-results/excel-table-design.png' });
  await command('Cancel');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'Reimport.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: output,
  });
  await expect(cell('E4')).toHaveText('40');
});

test('new workbooks normalize duplicate and blank table headers, reject overlap and export editable tables', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Excel', exact: true }).click();
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const command = (name: string) => page.getByRole('button', { name, exact: true }).click();
  for (const [ref, value] of [
    ['A1', 'Amount'],
    ['B1', 'Amount'],
    ['A2', '2'],
    ['B2', '3'],
  ]) {
    await cell(ref).click();
    await page.keyboard.type(value);
    await page.keyboard.press('Enter');
  }
  await cell('A1').click();
  await cell('C3').click({ modifiers: ['Shift'] });
  await command('Format as table');
  await page.getByLabel('Table name', { exact: true }).fill('Orders');
  await command('Create table');
  await expect(cell('B1')).toHaveText('Amount2');
  await expect(cell('C1')).toHaveText('Column3');
  await cell('D1').click();
  await page.keyboard.type('=SUM(Orders[Amount])');
  await page.keyboard.press('Enter');
  await expect(cell('D1')).toHaveText('2');
  await cell('D3').click();
  await command('Format as table');
  await page.getByLabel('Table range', { exact: true }).fill('C2:D4');
  await command('Create table');
  await expect(page.getByRole('alert')).toHaveText('Tables cannot overlap.');
  await command('Cancel');
  await command('Add sheet');
  await expect(page.getByRole('button', { name: 'Sheet 2', exact: true })).toBeVisible();
  await command('Sheet 1');
  const colors = [];
  for (const ref of ['A1', 'A2', 'A3'])
    colors.push({
      ref,
      ...(await cell(ref).evaluate((el) => ({
        fill: getComputedStyle(el).backgroundColor,
        color: getComputedStyle(el).color,
        bold: Number(getComputedStyle(el).fontWeight) >= 600,
      }))),
    });
  await fs.mkdir('.local/xlsx-tables', { recursive: true });
  await fs.writeFile('.local/xlsx-tables/new-browser-colors.json', JSON.stringify(colors));
  await command('Export');
  const pending = page.waitForEvent('download');
  await command('XLSX file Editable in Microsoft Excel');
  const bytes = await fs.readFile((await (await pending).path())!);
  await fs.mkdir('.local/xlsx-tables', { recursive: true });
  await fs.writeFile('.local/xlsx-tables/new.xlsx', bytes);
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'New reimport.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  await expect(cell('D1')).toHaveText('2');
  await cell('B2').click();
  await command('Table design');
  await expect(page.getByLabel('Table name', { exact: true })).toHaveValue('Orders');
});
