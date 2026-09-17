import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { tableRenameFixture } from './fixtures/xlsx-table-rename';

test('direct table header entry and rectangular paste retain column identities through undo, reload and native exports', async ({
  page,
}) => {
  const root = '.local/xlsx-header-entry';
  await fs.mkdir(root, { recursive: true });
  const input = Buffer.from(await tableRenameFixture());
  await fs.writeFile(`${root}/source.xlsx`, input);
  await page.goto('/');
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'Header entry.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: input,
  });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const command = (name: string) => page.getByRole('button', { name, exact: true }).click();
  const enter = async (ref: string, value: string) => {
    await cell(ref).click();
    await page.keyboard.type(value);
    await page.keyboard.press('Enter');
  };
  const headers = async (names: string[]) => {
    for (const [i, ref] of ['A1', 'B1', 'C1'].entries())
      await expect(cell(ref)).toHaveText(names[i]);
  };
  const paste = async (text: string) => {
    await cell('A1').click();
    await page.evaluate((text) => {
      const data = new DataTransfer();
      data.setData('text/plain', text);
      document.activeElement!.dispatchEvent(
        new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
      );
    }, text);
  };
  const download = async (stage: string) => {
    await command('Export');
    const pending = page.waitForEvent('download');
    await command('XLSX file Editable in Microsoft Excel');
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${stage}.xlsx`, bytes);
    return bytes;
  };
  await enter('B1', 'Units sold');
  await headers(['Region', 'Units sold', 'Price']);
  await expect(cell('E1')).toHaveText('10');
  await download('typed');
  await enter('B1', 'Price');
  await headers(['Region', 'Price', 'Price2']);
  await expect(cell('H4')).toHaveText('110');
  await download('duplicate');
  await command('Undo');
  await headers(['Region', 'Units sold', 'Price']);
  await command('Redo');
  await headers(['Region', 'Price', 'Price2']);
  await paste('Price\tPrice\tPrice2\nNorth\t4\t40');
  await headers(['Price', 'Price3', 'Price2']);
  await expect(cell('E1')).toHaveText('12');
  await expect(cell('C2')).toHaveText('40');
  await download('paste');
  await command('Undo');
  await headers(['Region', 'Price', 'Price2']);
  await expect(cell('B2')).toHaveText('2');
  await command('Redo');
  await headers(['Price', 'Price3', 'Price2']);
  await paste('Price3\tPrice\tPrice2');
  await headers(['Price3', 'Price', 'Price2']);
  await expect(cell('E1')).toHaveText('12');
  await download('swap');
  await cell('B1').click();
  await page.keyboard.press('Delete');
  await headers(['Price3', 'Column1', 'Price2']);
  await download('blank');
  await paste('\t\t');
  await headers(['Column1', 'Column2', 'Column3']);
  await download('clear');
  await enter('B1', "Net, [#@']");
  await headers(['Column1', "Net, [#@']", 'Column3']);
  await expect(cell('E1')).toHaveText('12');
  await download('escaped');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await command('Recent files');
  await command('Header entry');
  await headers(['Column1', "Net, [#@']", 'Column3']);
  await expect(cell('E1')).toHaveText('12');
  await enter('B1', 'X'.repeat(260));
  await headers(['Column1', 'X'.repeat(255), 'Column3']);
  await download('long');
  await enter('B1', "'=Heading");
  await headers(['Column1', '=Heading', 'Column3']);
  const output = await download('literal');
  await enter('B1', '=1+2');
  await expect(
    page.getByText(
      'Table headings are text. Start with an apostrophe to enter a heading beginning with =.',
      { exact: true },
    ),
  ).toBeVisible();
  await headers(['Column1', '=Heading', 'Column3']);
  const zip = await JSZip.loadAsync(output),
    original = await JSZip.loadAsync(input);
  for (const path of [
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/comments1.xml',
    'xl/worksheets/sheet2.xml',
    'custom/preservation.xml',
  ])
    expect(await zip.file(path)!.async('string')).toBe(await original.file(path)!.async('string'));
  await page.locator('input[type=file]').first().setInputFiles({
    name: 'Header output.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: output,
  });
  await headers(['Column1', '=Heading', 'Column3']);
  await expect(cell('E1')).toHaveText('12');
});

test('unsupported header dependencies reject the entire pasted header/data rectangle', async ({
  page,
}) => {
  const zip = await JSZip.loadAsync(await tableRenameFixture());
  zip.file('xl/connections.xml', '<connections/>');
  await page.goto('/');
  await page
    .locator('input[type=file]')
    .first()
    .setInputFiles({
      name: 'External headers.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: await zip.generateAsync({ type: 'nodebuffer' }),
    });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  await cell('A1').click();
  await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData('text/plain', 'Region\tUnits\tPrice\nNorth\t99\t10');
    document.activeElement!.dispatchEvent(
      new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
    );
  });
  await expect(
    page.getByText(
      'Table renaming with pivots, slicers, queries, external links, data models or macros is not supported yet.',
      { exact: true },
    ),
  ).toBeVisible();
  await expect(cell('B1')).toHaveText('Quantity');
  await expect(cell('B2')).toHaveText('2');
  await expect(cell('E1')).toHaveText('10');
});

test('an in-flight header repair prevents exporting stale content and resumes editing after it completes', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .locator('input[type=file]')
    .first()
    .setInputFiles({
      name: 'Delayed headers.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: Buffer.from(await tableRenameFixture()),
    });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  await expect(cell('B1')).toHaveText('Quantity');
  await page.evaluate(() => {
    const target = window as Window & { releaseHeader?: () => void; headerBlocked?: boolean };
    const original = Blob.prototype.arrayBuffer;
    Blob.prototype.arrayBuffer = async function () {
      Blob.prototype.arrayBuffer = original;
      const data = await original.call(this);
      target.headerBlocked = true;
      await new Promise<void>((resolve) => {
        target.releaseHeader = resolve;
      });
      return data;
    };
  });
  await cell('B1').click();
  await page.keyboard.type('Units');
  await page.keyboard.press('Enter');
  await expect
    .poll(() => page.evaluate(() => (window as Window & { headerBlocked?: boolean }).headerBlocked))
    .toBe(true);
  await expect(page.getByRole('button', { name: 'Export', exact: true })).toBeDisabled();
  await expect(page.locator('.sheet-editor')).toHaveAttribute(
    'aria-busy',
    'true',
  );
  await page.evaluate(() => (window as Window & { releaseHeader?: () => void }).releaseHeader!());
  await expect(cell('B1')).toHaveText('Units');
  await expect(page.getByRole('button', { name: 'Export', exact: true })).toBeEnabled();
  await cell('B2').click();
  await page.keyboard.type('6');
  await page.keyboard.press('Enter');
  await expect(cell('E1')).toHaveText('14');
});
