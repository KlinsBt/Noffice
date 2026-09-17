import { test, expect, type Page } from '@playwright/test';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import ExcelJS from 'exceljs';
import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import fs from 'node:fs/promises';

async function load(page: Page, name: string, buffer: Buffer) {
  await page.goto('/');
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles({ name, buffer, mimeType: 'application/octet-stream' });
}
async function download(page: Page, format: string) {
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page.getByRole('button', { name: new RegExp(`^${format} file`) }).click();
  return fs.readFile((await (await pending).path())!);
}
async function reload(page: Page, name: string) {
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click();
}

test('Word finds across formatted runs and replaces selected/all matches through undo, reload and retained DOCX export', async ({
  page,
}) => {
  const input = await Packer.toBuffer(
    new Document({
      sections: [
        {
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: 'Project ', bold: true }),
                new TextRun({ text: 'Atlas', italics: true }),
                new TextRun(' launch'),
              ],
            }),
            new Paragraph('project atlas and Project Atlases'),
          ],
        },
      ],
    }),
  );
  await load(page, 'Search document.docx', input);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.click();
  await page.keyboard.press('Control+h');
  await page.getByRole('textbox', { name: 'Find text', exact: true }).fill('Project Atlas');
  await page.getByRole('checkbox', { name: 'Whole words only', exact: true }).check();
  await page.getByRole('checkbox', { name: 'Match case', exact: true }).check();
  await page.getByRole('button', { name: 'Find next', exact: true }).click();
  await expect(page.locator('.search-panel [role=status]')).toHaveText('1 of 1 matches');
  expect(await page.evaluate(() => window.getSelection()?.toString())).toBe('Project Atlas');
  await page.getByRole('textbox', { name: 'Replacement text', exact: true }).fill('Plan $&');
  await page.getByRole('button', { name: 'Replace', exact: true }).click();
  await expect(editor.locator('p').first()).toHaveText('Plan $& launch');
  await expect(editor.locator('p').first().locator('strong')).toHaveText('Plan $&');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor.locator('p').first()).toHaveText('Project Atlas launch');
  await page.getByRole('checkbox', { name: 'Match case', exact: true }).uncheck();
  await page.getByRole('button', { name: 'Replace all', exact: true }).click();
  await expect(editor.locator('p').last()).toHaveText('Plan $& and Project Atlases');
  await reload(page, 'Search document');
  await expect(editor.locator('p').first()).toHaveText('Plan $& launch');
  const output = await download(page, 'DOCX'),
    before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'word/document.xml')
      expect(await after.file(path)!.async('nodebuffer'), path).toEqual(
        await before.file(path)!.async('nodebuffer'),
      );
  await load(page, 'Search result.docx', output);
  await expect(editor.locator('p').last()).toHaveText('Plan $& and Project Atlases');
});

test('Excel searches values and replaces workbook formulas/text while retaining wrap and vertical alignment on export', async ({
  page,
}) => {
  const workbook = new ExcelJS.Workbook(),
    first = workbook.addWorksheet('First'),
    second = workbook.addWorksheet('Second');
  first.getCell('A1').value = 'Atlas';
  first.getCell('A1').font = { bold: true };
  first.getCell('B1').value = { formula: 'IF(A1="Atlas",10,0)', result: 10 };
  first.getCell('A2').value = 'Atlas';
  second.getCell('A1').value = 'Atlas';
  await load(page, 'Search book.xlsx', Buffer.from(await workbook.xlsx.writeBuffer()));
  await page.getByRole('region', { name: 'Spreadsheet grid' }).focus();
  await page.keyboard.press('Control+h');
  await page.getByRole('textbox', { name: 'Find cell value', exact: true }).fill('Atlas');
  await page.getByRole('textbox', { name: 'Replacement text', exact: true }).fill('Orion');
  await page.getByRole('combobox', { name: 'Search within', exact: true }).selectOption('workbook');
  await page.getByRole('button', { name: 'Replace all', exact: true }).click();
  await expect(page.locator('.search-panel [role=status]')).toHaveText('Replaced 4 cells.');
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveText('Orion');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveText('Atlas');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await page.getByRole('combobox', { name: 'Look in', exact: true }).selectOption('values');
  await page.getByRole('textbox', { name: 'Find cell value', exact: true }).fill('10');
  await page.getByRole('button', { name: 'Find next', exact: true }).click();
  await expect(page.getByRole('textbox', { name: 'Formula bar', exact: true })).toHaveValue(
    '=IF(A1="Orion",10,0)',
  );
  await page.getByRole('button', { name: 'Replace all', exact: true }).click();
  await expect(page.locator('.search-panel [role=status]')).toContainText('Choose Formulas');
  await page.getByRole('button', { name: 'Close find and replace', exact: true }).click();
  await page.getByRole('gridcell', { name: 'A1', exact: true }).click();
  await page.getByRole('gridcell', { name: 'B2', exact: true }).click({ modifiers: ['Shift'] });
  await page.getByRole('button', { name: 'Wrap text', exact: true }).click();
  await page
    .getByRole('combobox', { name: 'Vertical alignment', exact: true })
    .selectOption('middle');
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveCSS(
    'vertical-align',
    'middle',
  );
  await reload(page, 'Search book');
  const output = await download(page, 'XLSX'),
    read = new ExcelJS.Workbook();
  await read.xlsx.load(output as never);
  expect(read.getWorksheet('First')!.getCell('A1').value).toBe('Orion');
  expect(read.getWorksheet('Second')!.getCell('A1').value).toBe('Orion');
  expect(read.getWorksheet('First')!.getCell('B1').formula).toBe('IF(A1="Orion",10,0)');
  expect(read.getWorksheet('First')!.getCell('A1').alignment).toMatchObject({
    wrapText: true,
    vertical: 'middle',
  });
  expect(read.getWorksheet('First')!.getCell('A1').font.bold).toBe(true);
});

test('Excel wildcards and escaped question marks respect workbook protection without partial changes', async ({
  page,
}) => {
  const workbook = new ExcelJS.Workbook(),
    first = workbook.addWorksheet('Visible'),
    locked = workbook.addWorksheet('Locked');
  first.getCell('A1').value = 'west-2024';
  first.getCell('A2').value = 'east-2025';
  first.getCell('A3').value = 'west?';
  locked.getCell('A1').value = 'west-2026';
  await locked.protect('test', { spinCount: 1 });
  await load(page, 'Protected search.xlsx', Buffer.from(await workbook.xlsx.writeBuffer()));
  await page.getByRole('button', { name: 'Find cells', exact: true }).click();
  await page.getByRole('textbox', { name: 'Find cell value', exact: true }).fill('west*');
  await page.getByRole('textbox', { name: 'Replacement text', exact: true }).fill('Updated');
  await page.getByRole('combobox', { name: 'Search within', exact: true }).selectOption('workbook');
  await page.getByRole('button', { name: 'Replace all', exact: true }).click();
  await expect(page.locator('.search-panel [role=status]')).toContainText('No cells were changed');
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveText('west-2024');
  await page.getByRole('combobox', { name: 'Search within', exact: true }).selectOption('sheet');
  await page.getByRole('textbox', { name: 'Find cell value', exact: true }).fill('west~?');
  await page.getByRole('button', { name: 'Replace all', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'A3', exact: true })).toHaveText('Updated');
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveText('west-2024');
  await page.screenshot({ path: 'test-results/search-excel-desktop.png' });
  await page.setViewportSize({ width: 390, height: 900 });
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(390);
  await page.getByRole('textbox', { name: 'Find cell value', exact: true }).press('Escape');
  await expect(page.locator('.search-panel')).toHaveCount(0);
});

test('PowerPoint navigates occurrences across slides and retains replacement through undo, reload and PPTX export', async ({
  page,
}) => {
  const pptx = new PptxGenJS();
  pptx.addSlide().addText('Atlas Atlas', { x: 1, y: 1, w: 6, h: 1, bold: true });
  pptx.addSlide().addText('Atlas launch', { x: 1, y: 1, w: 6, h: 1, italic: true });
  const input = Buffer.from((await pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer);
  await load(page, 'Search deck.pptx', input);
  await page.getByRole('application', { name: 'Presentation editor', exact: true }).focus();
  await page.keyboard.press('Control+h');
  await page.getByRole('textbox', { name: 'Find text', exact: true }).fill('Atlas');
  await page.getByRole('button', { name: 'Find previous', exact: true }).click();
  await expect(page.locator('.search-panel [role=status]')).toHaveText('3 of 3 matches · slide 2');
  await page.getByRole('textbox', { name: 'Replacement text', exact: true }).fill('Orion');
  await page.getByRole('button', { name: 'Replace', exact: true }).click();
  await expect(page.getByRole('textbox', { name: 'Object text', exact: true })).toHaveValue(
    'Orion launch',
  );
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByRole('textbox', { name: 'Object text', exact: true })).toHaveValue(
    'Atlas launch',
  );
  await page.getByRole('button', { name: 'Replace all', exact: true }).click();
  await expect(page.locator('.search-panel [role=status]')).toHaveText('Replaced 3 matches.');
  await reload(page, 'Search deck');
  const output = await download(page, 'PPTX'),
    before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && !/^ppt\/slides\/slide\d+\.xml$/.test(path))
      expect(await after.file(path)!.async('nodebuffer'), path).toEqual(
        await before.file(path)!.async('nodebuffer'),
      );
  expect(await after.file('ppt/slides/slide1.xml')!.async('string')).toContain('Orion Orion');
  expect(await after.file('ppt/slides/slide2.xml')!.async('string')).toContain('Orion launch');
  await load(page, 'Search deck result.pptx', output);
  await expect(page.locator('.slide-element').first()).toContainText('Orion Orion');
});
