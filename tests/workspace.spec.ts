import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';

async function openHome(page: Page) {
  await page.goto('/');
  await expect(page.getByRole('heading', { name: 'What will you create today?' })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Start Word', exact: true })).toBeEnabled();
}
async function openSample(page: Page, name: string) {
  await openHome(page);
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click();
}
async function saved(page: Page) {
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
}
async function exportDownload(page: Page, name: string) {
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const result = page.waitForEvent('download');
  await page.getByRole('button', { name }).click();
  return result;
}

test('mode landing, recent files, search, templates, favorites and recoverable trash', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await openHome(page);
  await page.screenshot({ path: 'test-results/home-desktop.png', fullPage: true });
  await expect(page.locator('.sidebar')).toHaveCount(0);
  await expect(page.locator('.mode-card')).toHaveCount(3);
  await expect(page.locator('.file-row')).toHaveCount(0);
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('textbox', { name: 'Search files' }).fill('budget');
  await expect(page.locator('.file-row')).toHaveCount(1);
  await page.getByRole('textbox', { name: 'Search files' }).fill('');
  await page.getByRole('button', { name: 'More options for Project budget' }).click();
  await page.getByRole('button', { name: 'Add to starred' }).click();
  await page.getByRole('button', { name: 'Starred', exact: true }).click();
  await expect(page.locator('.file-row')).toHaveCount(1);
  await page.getByRole('button', { name: 'More options for Project budget' }).click();
  await page.getByRole('button', { name: 'Move to Trash' }).click();
  await page.getByRole('button', { name: 'Trash', exact: true }).click();
  await expect(page.locator('.file-row')).toHaveCount(1);
  await page.getByRole('button', { name: 'More options for Project budget' }).click();
  await page.getByRole('button', { name: 'Restore file' }).click();
  await expect(page.locator('.file-row')).toHaveCount(0);
  await page.getByRole('button', { name: 'Create new', exact: true }).click();
  await page.getByRole('button', { name: 'Spreadsheet template', exact: true }).click();
  await expect(page.getByRole('region', { name: 'Spreadsheet grid' })).toBeVisible();
  expect(errors).toEqual([]);
});

test('Write: editing, formatting, find and replace, table, autosave and DOCX export', async ({
  page,
}) => {
  await openSample(page, 'A fresh start');
  const editor = page.getByRole('textbox', { name: 'Document text' });
  await editor.click();
  await page.keyboard.press('Control+End');
  await page.keyboard.press('Enter');
  await page.keyboard.type('A browser office that belongs to you.');
  await page.getByRole('button', { name: 'Find & replace' }).click();
  await page.getByRole('textbox', { name: 'Find text', exact: true }).fill('browser office');
  await page.getByRole('textbox', { name: 'Replacement text' }).fill('private workspace');
  await page.getByRole('button', { name: 'Replace all' }).click();
  await expect(editor).toContainText('A private workspace that belongs to you.');
  await editor.click();
  await page.keyboard.press('Control+End');
  await page.keyboard.press('Enter');
  await page.getByRole('button', { name: 'Bold', exact: true }).click();
  await page.keyboard.type('Bold evidence');
  await expect(editor.locator('strong')).toContainText('Bold evidence');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Table', exact: true }).click();
  await expect(editor.locator('table')).toHaveCount(1);
  await saved(page);
  await page.screenshot({ path: 'test-results/write-desktop.png' });
  const download = await exportDownload(page, 'DOCX file Editable in Microsoft Word');
  const bytes = await fs.readFile((await download.path())!);
  const zip = await JSZip.loadAsync(bytes);
  const document = await zip.file('word/document.xml')!.async('string');
  expect(document).toContain('Bold evidence');
  expect(document).toContain('w:tbl');
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'A fresh start', exact: true }).click();
  await expect(page.getByRole('textbox', { name: 'Document text' })).toContainText(
    'private workspace',
  );
});

test('Calculate: entry, formula recalculation, range fill, formatting, sheets and XLSX export', async ({
  page,
}) => {
  await openSample(page, 'Project budget');
  await expect(page.getByRole('gridcell', { name: 'D8', exact: true })).toHaveText('$8,980.00');
  await page.getByRole('gridcell', { name: 'B3', exact: true }).dblclick();
  await page.getByRole('textbox', { name: 'Edit B3', exact: true }).fill('20');
  await page.keyboard.press('Enter');
  await expect(page.getByRole('gridcell', { name: 'D3', exact: true })).toHaveText('$1,700.00');
  await page.getByRole('gridcell', { name: 'F1', exact: true }).click();
  await page.getByRole('textbox', { name: 'Formula bar' }).fill('=SUM(D3:D6)');
  await page.keyboard.press('Enter');
  await expect(page.getByRole('gridcell', { name: 'F1', exact: true })).toHaveText('9660');
  await page.getByRole('gridcell', { name: 'F2', exact: true }).click({ modifiers: ['Shift'] });
  await page.getByRole('button', { name: 'Fill down', exact: true }).click();
  await page.getByRole('gridcell', { name: 'F2', exact: true }).click();
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('=SUM(D4:D7)');
  await page.getByRole('button', { name: 'Bold', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'F2', exact: true })).toHaveCSS(
    'font-weight',
    '650',
  );
  await page.getByRole('button', { name: 'Add sheet' }).click();
  await expect(page.getByRole('button', { name: 'Sheet 2', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Budget', exact: true }).click();
  await saved(page);
  await page.screenshot({ path: 'test-results/calculate-desktop.png' });
  const download = await exportDownload(page, 'XLSX file Editable in Microsoft Excel');
  const zip = await JSZip.loadAsync(await fs.readFile((await download.path())!));
  expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toContain(
    '<f>SUM(D3:D6)</f>',
  );
  expect(zip.file('xl/worksheets/sheet2.xml')).not.toBeNull();
});

test('Present: edit, add, drag, reorder, notes, slideshow keyboard and PPTX export', async ({
  page,
}) => {
  await openSample(page, 'The next chapter');
  await page.getByRole('button', { name: 'Text box', exact: true }).click();
  await page
    .getByRole('textbox', { name: 'Object text', exact: true })
    .fill('A tested presentation');
  const el = page.locator('.slide-stage .slide-element.selected');
  const box = (await el.boundingBox())!;
  await page.mouse.move(box.x + 20, box.y + 20);
  await page.mouse.down();
  await page.mouse.move(box.x + 60, box.y + 50, { steps: 4 });
  await page.mouse.up();
  await expect(page.getByRole('spinbutton', { name: 'Object x', exact: true })).not.toHaveValue(
    '100',
  );
  await page
    .getByRole('textbox', { name: 'Speaker notes' })
    .fill('Remember to thank the audience.');
  await page.getByRole('button', { name: 'Duplicate slide', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Slide 4', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Move slide down', exact: true }).click();
  await page.screenshot({ path: 'test-results/present-desktop.png' });
  await page.getByRole('button', { name: 'Present', exact: true }).last().click();
  await expect(page.getByRole('dialog', { name: 'Slideshow' })).toBeVisible();
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Escape');
  await expect(page.getByRole('dialog', { name: 'Slideshow' })).not.toBeVisible();
  await saved(page);
  const download = await exportDownload(page, 'PPTX file Editable in Microsoft PowerPoint');
  const zip = await JSZip.loadAsync(await fs.readFile((await download.path())!));
  expect(zip.file('ppt/slides/slide4.xml')).not.toBeNull();
  expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toContain(
    'A tested presentation',
  );
  expect(await zip.file('ppt/notesSlides/notesSlide1.xml')!.async('string')).toContain(
    'Remember to thank the audience.',
  );
});

test('app switching preserves unsaved content and native backups restore it', async ({ page }) => {
  await openSample(page, 'A fresh start');
  const editor = page.getByRole('textbox', { name: 'Document text' });
  await editor.click();
  await page.keyboard.press('Control+End');
  await page.keyboard.type(' Switching survives.');
  await page.locator('.app-switcher').getByRole('button', { name: 'Excel', exact: true }).click();
  await expect(page.getByRole('region', { name: 'Spreadsheet grid' })).toBeVisible();
  await page
    .locator('.app-switcher')
    .getByRole('button', { name: 'PowerPoint', exact: true })
    .click();
  await expect(page.getByLabel('Slide canvas', { exact: true })).toBeVisible();
  await page.locator('.app-switcher').getByRole('button', { name: 'Word', exact: true }).click();
  await expect(editor).toContainText('Switching survives.');
  await saved(page);
  const download = await exportDownload(
    page,
    'Noffice backup Full editable model and retained original · .noffice',
  );
  const path = (await download.path())!;
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'restored.noffice',
    mimeType: 'application/json',
    buffer: await fs.readFile(path),
  });
  await expect(editor).toContainText('Switching survives.');
});

test('imports DOCX, XLSX and PPTX generated by the suite and retains originals', async ({
  page,
}) => {
  for (const [name, button, extension] of [
    ['A fresh start', 'DOCX file Editable in Microsoft Word', 'docx'],
    ['Project budget', 'XLSX file Editable in Microsoft Excel', 'xlsx'],
    ['The next chapter', 'PPTX file Editable in Microsoft PowerPoint', 'pptx'],
  ]) {
    await openSample(page, name);
    const download = await exportDownload(page, button);
    const buffer = await fs.readFile((await download.path())!);
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `roundtrip.${extension}`,
      mimeType: 'application/octet-stream',
      buffer,
    });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'roundtrip',
    );
    await expect(page.getByRole('button', { name: /Conversion notes/ })).toBeVisible();
    await saved(page);
    await page.getByRole('button', { name: /Conversion notes/ }).click();
    const originalDownload = page.waitForEvent('download');
    await page.getByRole('button', { name: 'Download unchanged original' }).click();
    const original = await originalDownload;
    expect(await fs.readFile((await original.path())!)).toEqual(buffer);
  }
});

test('malicious HTML cannot load remote resources or execute scripts', async ({ page }) => {
  await openHome(page);
  const remote: string[] = [];
  page.on('request', (request) => {
    if (!request.url().startsWith('http://127.0.0.1:4183')) remote.push(request.url());
  });
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'unsafe.html',
    mimeType: 'text/html',
    buffer: Buffer.from(
      '<h1>Safe heading</h1><script>window.pwned=true</script><img src="https://example.com/tracker"><p onclick="window.pwned=true">Safe text</p><p style="background-image:url(https://example.com/css)">No fetch</p>',
    ),
  });
  await expect(page.getByRole('textbox', { name: 'Document text' })).toContainText('Safe heading');
  expect(
    await page.evaluate(() => (window as unknown as { pwned?: boolean }).pwned),
  ).toBeUndefined();
  expect(remote).toEqual([]);
});

test('CSV paste, range copy, sorting and undo operate on actual cells', async ({ page }) => {
  await openHome(page);
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'data.csv',
    mimeType: 'text/csv',
    buffer: Buffer.from('3,three\n1,one\n2,two'),
  });
  await page.getByRole('gridcell', { name: 'A1', exact: true }).click();
  await page.getByRole('gridcell', { name: 'B3', exact: true }).click({ modifiers: ['Shift'] });
  await page.getByRole('button', { name: 'Sort ascending' }).click();
  await expect(page.getByRole('gridcell', { name: 'B1', exact: true })).toHaveText('one');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'B1', exact: true })).toHaveText('three');
  await page.getByRole('gridcell', { name: 'D1', exact: true }).click();
  await page.getByRole('region', { name: 'Spreadsheet grid' }).evaluate((el) => {
    const data = new DataTransfer();
    data.setData('text/plain', '4\t5\n6\t=SUM(D1:E1)');
    el.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true }));
  });
  await expect(page.getByRole('gridcell', { name: 'E2', exact: true })).toHaveText('9');
});

test('stale browser tab cannot overwrite a more recent document', async ({ page, context }) => {
  await openSample(page, 'A fresh start');
  const second = await context.newPage();
  await openSample(second, 'A fresh start');
  await page.getByRole('textbox', { name: 'File name', exact: true }).fill('First tab wins');
  await saved(page);
  await second.getByRole('textbox', { name: 'File name', exact: true }).fill('Stale overwrite');
  await expect(second.getByRole('alert')).toContainText('another tab');
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await expect(page.getByRole('button', { name: 'First tab wins', exact: true })).toBeVisible();
});

test('local history restores an earlier saved document', async ({ page }) => {
  await openSample(page, 'A fresh start');
  await page.getByRole('textbox', { name: 'File name', exact: true }).fill('Renamed document');
  await saved(page);
  await page.getByRole('button', { name: 'Version history', exact: true }).click();
  await page.locator('.version-list>div').last().getByRole('button', { name: 'Restore' }).click();
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'A fresh start',
  );
});

test('mobile workspace fits its viewport and supports file creation', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await openHome(page);
  await page.screenshot({ path: 'test-results/home-mobile.png', fullPage: true });
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
  await page.getByRole('button', { name: 'Start Word', exact: true }).click();
  await expect(page.getByRole('textbox', { name: 'Document text' })).toBeVisible();
});

test('static app reopens offline and exports without external requests', async ({
  page,
  context,
}) => {
  await openHome(page);
  await page.evaluate(async () => {
    await navigator.serviceWorker.ready;
  });
  await expect
    .poll(() => page.evaluate(() => Boolean(navigator.serviceWorker.controller)))
    .toBe(true);
  await context.setOffline(true);
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await expect(page.getByRole('button', { name: 'A fresh start', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Project budget', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'D8', exact: true })).toHaveText('$8,980.00');
  const download = await exportDownload(page, 'XLSX file Editable in Microsoft Excel');
  expect((await fs.stat((await download.path())!)).size).toBeGreaterThan(1000);
});

test('local images export into DOCX and PPTX packages', async ({ page }) => {
  await openSample(page, 'A fresh start');
  const data = await page.evaluate(() => {
    const canvas = document.createElement('canvas');
    canvas.width = 80;
    canvas.height = 40;
    const ctx = canvas.getContext('2d')!;
    ctx.fillStyle = '#418369';
    ctx.fillRect(0, 0, 80, 40);
    return canvas.toDataURL('image/png').split(',')[1];
  });
  const file = {
    name: 'local-image.png',
    mimeType: 'image/png',
    buffer: Buffer.from(data, 'base64'),
  };
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.locator('.word-editor input[type=file]').setInputFiles(file);
  // ProseMirror adds an empty separator image beside inline non-editable nodes.
  const insertedImages = page
    .getByRole('textbox', { name: 'Document text' })
    .locator('img:not(.ProseMirror-separator)');
  await expect(insertedImages).toHaveCount(1);
  await expect(insertedImages).toHaveAttribute('alt', 'local-image.png');
  await expect(insertedImages).toHaveAttribute('src', `data:image/png;base64,${data}`);
  const doc = await exportDownload(page, 'DOCX file Editable in Microsoft Word');
  const docZip = await JSZip.loadAsync(await fs.readFile((await doc.path())!));
  expect(
    Object.keys(docZip.files).some(
      (name) => name.startsWith('word/media/') && name.endsWith('.png'),
    ),
  ).toBe(true);
  await page
    .locator('.app-switcher')
    .getByRole('button', { name: 'PowerPoint', exact: true })
    .click();
  await page.locator('.slide-editor input[type=file]').setInputFiles(file);
  await expect(page.locator('.slide-stage img')).toHaveCount(1);
  const deck = await exportDownload(page, 'PPTX file Editable in Microsoft PowerPoint');
  const deckZip = await JSZip.loadAsync(await fs.readFile((await deck.path())!));
  expect(
    Object.keys(deckZip.files).some(
      (name) => name.startsWith('ppt/media/') && name.endsWith('.png'),
    ),
  ).toBe(true);
});
