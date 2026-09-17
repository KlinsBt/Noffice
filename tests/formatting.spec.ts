import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';

async function start(page: Page, mode: string) {
  await page.goto('/');
  await page.getByRole('button', { name: `Start ${mode}`, exact: true }).click();
}
async function officeXML(page: Page, type: string, entry: string) {
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page.getByRole('button', { name: new RegExp(`^${type} file`) }).click();
  const download = await pending;
  const zip = await JSZip.loadAsync(await fs.readFile((await download.path())!));
  return { xml: await zip.file(entry)!.async('string'), path: (await download.path())! };
}
async function saved(page: Page) {
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
}

test('Word font, point size, script, spacing, orientation and selection state persist and export', async ({
  page,
}) => {
  await start(page, 'Word');
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.fill('Typography evidence');
  await page.keyboard.press('Control+a');
  await page
    .getByRole('combobox', { name: 'Font family', exact: true })
    .selectOption('Source Serif 4');
  await page.getByRole('spinbutton', { name: 'Font size', exact: true }).fill('24');
  await page.getByRole('spinbutton', { name: 'Font size', exact: true }).press('Tab');
  await page.getByRole('button', { name: 'Superscript', exact: true }).click();
  await page.getByRole('combobox', { name: 'Line spacing', exact: true }).selectOption('2');
  await expect(editor.locator('sup')).toHaveText('Typography evidence');
  // Whole-paragraph font editing now also preserves its 24pt paragraph mark.
  // Vertical effects retain the CSS fallback; native superscript metrics remain open.
  await expect(editor.locator('p')).toHaveCSS('font-size', '32px');
  await expect(editor.locator('p')).toHaveCSS('line-height', '64px');
  await expect(editor.locator('span').first()).toHaveCSS('font-size', '32px');
  await expect(editor.locator('span').first()).toHaveCSS('font-family', '"Source Serif 4"');
  await expect(page.getByRole('combobox', { name: 'Font family', exact: true })).toHaveValue(
    'Source Serif 4',
  );
  await page.getByRole('button', { name: 'Layout', exact: true }).click();
  await page
    .getByRole('combobox', { name: 'Page orientation', exact: true })
    .selectOption('landscape');
  await expect(page.locator('.paper')).toHaveClass(/landscape/);
  await saved(page);
  const { xml } = await officeXML(page, 'DOCX', 'word/document.xml');
  expect(xml).toContain('Source Serif 4');
  expect(xml).toContain('w:sz w:val="48"');
  expect(xml).toContain('w:vertAlign w:val="superscript"');
  expect(xml).toContain('w:orient="landscape"');
  expect(xml).toContain('w:line="480"');
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Untitled document', exact: true }).click();
  await expect(editor.locator('sup')).toHaveText('Typography evidence');
  await expect(page.locator('.paper')).toHaveClass(/landscape/);
  await editor.click();
  await page.keyboard.press('Control+a');
  await expect(page.getByRole('combobox', { name: 'Font family', exact: true })).toHaveValue(
    'Source Serif 4',
  );
  await expect(page.getByRole('spinbutton', { name: 'Font size', exact: true })).toHaveValue('24');
  await page.getByRole('button', { name: 'Subscript', exact: true }).click();
  await expect(editor.locator('sub')).toHaveText('Typography evidence');
  await expect(editor.locator('sup')).toHaveCount(0);
  const exported = await officeXML(page, 'DOCX', 'word/document.xml');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'font-roundtrip.docx',
    mimeType: 'application/octet-stream',
    buffer: await fs.readFile(exported.path),
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'font-roundtrip',
  );
  await expect(editor.locator('sub')).toHaveText('Typography evidence');
  await expect(editor.locator('span').first()).toHaveCSS('font-family', '"Source Serif 4"');
  await expect(editor.locator('span').first()).toHaveCSS('font-size', '32px');
});

test('Word table column deletion, header toggling, merge, split and merged DOCX export', async ({
  page,
}) => {
  await start(page, 'Word');
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.click();
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Table', exact: true }).click();
  await expect(editor.locator('th')).toHaveCount(3);
  await editor.locator('th').first().click();
  await page.getByRole('button', { name: 'Header row', exact: true }).click();
  await expect(editor.locator('th')).toHaveCount(0);
  await page.getByRole('button', { name: 'Delete column', exact: true }).click();
  await expect(editor.locator('td')).toHaveCount(6);
  await editor.locator('td').nth(0).click();
  await editor
    .locator('td')
    .nth(1)
    .click({ modifiers: ['Shift'] });
  await page.getByRole('button', { name: 'Merge cells', exact: true }).click();
  await expect(editor.locator('td[colspan="2"]')).toHaveCount(1);
  const { xml } = await officeXML(page, 'DOCX', 'word/document.xml');
  expect(xml).toContain('w:gridSpan w:val="2"');
  await page.getByRole('button', { name: 'Split cell', exact: true }).click();
  await expect(editor.locator('td')).toHaveCount(6);
  await expect(editor.locator('td[colspan="2"]')).toHaveCount(0);
  await saved(page);
});

test('Excel applies font and underline to a range, undoes and preserves XLSX typography', async ({
  page,
}) => {
  await start(page, 'Excel');
  const a = page.getByRole('gridcell', { name: 'A1', exact: true });
  await a.dblclick();
  await page.getByRole('textbox', { name: 'Edit A1', exact: true }).fill('Font evidence');
  await page.keyboard.press('Enter');
  await a.click();
  await page.getByRole('gridcell', { name: 'B2', exact: true }).click({ modifiers: ['Shift'] });
  await page
    .getByRole('combobox', { name: 'Font family', exact: true })
    .selectOption('JetBrains Mono');
  await page.getByRole('spinbutton', { name: 'Font size', exact: true }).fill('18');
  await page.getByRole('spinbutton', { name: 'Font size', exact: true }).press('Tab');
  await page.getByRole('button', { name: 'Underline', exact: true }).click();
  await expect(a).toHaveCSS('font-family', /JetBrains Mono/);
  await expect(a).toHaveCSS('font-size', '24px');
  await expect(a).toHaveCSS('text-decoration-line', 'underline');
  await expect(page.getByRole('gridcell', { name: 'B2', exact: true })).toHaveCSS(
    'font-size',
    '24px',
  );
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(a).toHaveCSS('text-decoration-line', 'none');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  const { xml, path } = await officeXML(page, 'XLSX', 'xl/styles.xml');
  expect(xml).toContain('JetBrains Mono');
  expect(xml).toContain('<sz val="18"');
  expect(xml).toContain('<u/>');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'font-roundtrip.xlsx',
    mimeType: 'application/octet-stream',
    buffer: await fs.readFile(path),
  });
  await expect(a).toHaveText('Font evidence');
  await expect(a).toHaveCSS('font-family', /JetBrains Mono/);
  await expect(a).toHaveCSS('font-size', '24px');
  await expect(a).toHaveCSS('text-decoration-line', 'underline');
});

test('PowerPoint font, italic and underline appear in the slide, slideshow and PPTX roundtrip', async ({
  page,
}) => {
  await start(page, 'PowerPoint');
  await page.locator('.slide-element').first().click();
  await page
    .getByRole('combobox', { name: 'Font family', exact: true })
    .selectOption('Source Serif 4');
  await page.getByRole('button', { name: 'Italic text', exact: true }).click();
  await page.getByRole('button', { name: 'Underline text', exact: true }).click();
  await expect(page.locator('.slide-element').first()).toHaveCSS('font-family', /Source Serif 4/);
  await expect(page.locator('.slide-element').first()).toHaveCSS('font-style', 'italic');
  await page.getByRole('button', { name: 'Present', exact: true }).click();
  const preview = page
    .getByRole('dialog', { name: 'Slideshow' })
    .locator('foreignObject div')
    .first();
  await expect(preview).toHaveCSS('font-family', /Source Serif 4/);
  await expect(preview).toHaveCSS('text-decoration-line', 'underline');
  await page.keyboard.press('Escape');
  const { xml, path } = await officeXML(page, 'PPTX', 'ppt/slides/slide1.xml');
  expect(xml).toContain('typeface="Source Serif 4"');
  expect(xml).toContain('i="1"');
  expect(xml).toContain('u="sng"');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'font-roundtrip.pptx',
    mimeType: 'application/octet-stream',
    buffer: await fs.readFile(path),
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'font-roundtrip',
  );
  await expect(page.locator('.slide-element').first()).toHaveCSS('font-family', /Source Serif 4/);
  await expect(page.locator('.slide-element').first()).toHaveCSS('font-style', 'italic');
});

test('bundled fonts load offline and a local font persists and is available in every mode', async ({
  page,
  context,
}) => {
  const external: string[] = [];
  page.on('request', (request) => {
    if (!request.url().startsWith('http://127.0.0.1:4183') && !request.url().startsWith('data:'))
      external.push(request.url());
  });
  await start(page, 'Word');
  await page.getByRole('button', { name: 'Add font', exact: true }).click();
  await page.getByRole('textbox', { name: 'Imported font family name' }).fill('Local Test Serif');
  await page
    .getByLabel('Font file', { exact: true })
    .setInputFiles(
      'node_modules/@fontsource-variable/source-serif-4/files/source-serif-4-latin-wght-normal.woff2',
    );
  await page.getByRole('button', { name: 'Add font to this device', exact: true }).click();
  await expect(page.getByRole('combobox', { name: 'Font family', exact: true })).toHaveValue(
    'Local Test Serif',
  );
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.fill('Persistent font');
  await page.keyboard.press('Control+a');
  await page
    .getByRole('combobox', { name: 'Font family', exact: true })
    .selectOption('Local Test Serif');
  await saved(page);
  await page.evaluate(async () => {
    await navigator.serviceWorker.ready;
  });
  await expect.poll(() => page.evaluate(() => !!navigator.serviceWorker.controller)).toBe(true);
  await context.setOffline(true);
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Untitled document', exact: true }).click();
  await expect(editor.getByText('Persistent font', { exact: true })).toHaveCSS(
    'font-family',
    '"Local Test Serif"',
  );
  await expect
    .poll(() =>
      page.evaluate(() =>
        [...document.fonts].some(
          (face) => face.family === 'Local Test Serif' && face.status === 'loaded',
        ),
      ),
    )
    .toBe(true);
  const fonts = await page.evaluate(async () => {
    const names = ['Inter', 'Source Serif 4', 'JetBrains Mono'];
    return Promise.all(
      names.map(async (name) => (await document.fonts.load(`16px "${name}"`)).length),
    );
  });
  expect(fonts).toEqual([1, 1, 1]);
  await page.locator('.app-switcher').getByRole('button', { name: 'Excel', exact: true }).click();
  await page
    .getByRole('combobox', { name: 'Font family', exact: true })
    .selectOption('Local Test Serif');
  await expect(page.getByRole('gridcell', { name: 'A1', exact: true })).toHaveCSS(
    'font-family',
    /Local Test Serif/,
  );
  await page
    .locator('.app-switcher')
    .getByRole('button', { name: 'PowerPoint', exact: true })
    .click();
  await page.locator('.slide-element').first().click();
  await page
    .getByRole('combobox', { name: 'Font family', exact: true })
    .selectOption('Local Test Serif');
  await expect(page.locator('.slide-element').first()).toHaveCSS('font-family', /Local Test Serif/);
  expect(external).toEqual([]);
});

test('invalid local font reports an error without adding a family', async ({ page }) => {
  await start(page, 'Word');
  await page.getByRole('button', { name: 'Add font', exact: true }).click();
  await page.getByRole('textbox', { name: 'Imported font family name' }).fill('Broken Font');
  await page.getByLabel('Font file', { exact: true }).setInputFiles({
    name: 'broken.woff2',
    mimeType: 'font/woff2',
    buffer: Buffer.from('invalid font'),
  });
  await page.getByRole('button', { name: 'Add font to this device', exact: true }).click();
  await expect(page.getByRole('alert')).toBeVisible();
  await page.getByRole('button', { name: 'Close dialog', exact: true }).click();
  await expect(
    page
      .getByRole('combobox', { name: 'Font family', exact: true })
      .locator('option[value="Broken Font"]'),
  ).toHaveCount(0);
});
