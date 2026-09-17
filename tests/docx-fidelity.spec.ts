import { test, expect, type Page } from '@playwright/test';
import { Document, Header, Packer, Paragraph, TextRun, FootnoteReferenceRun } from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';

async function fixture() {
  const data = await Packer.toBuffer(
    new Document({
      footnotes: { 2: { children: [new Paragraph('Original footnote')] } },
      sections: [
        {
          headers: { default: new Header({ children: [new Paragraph('Retained header')] }) },
          properties: { page: { margin: { left: 1320, right: 1440, top: 1000, bottom: 1200 } } },
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: 'Alpha beta', bold: true, font: 'Georgia', size: 28 }),
              ],
            }),
            new Paragraph({
              children: [new TextRun('Referenced text'), new FootnoteReferenceRun(2)],
            }),
            new Paragraph('Final body paragraph'),
          ],
        },
      ],
    }),
  );
  const zip = await JSZip.loadAsync(data);
  zip.file('customXml/retained.xml', '<custom>Untouched content</custom>');
  return zip.generateAsync({ type: 'nodebuffer' });
}
async function open(page: Page, buffer: Buffer, name = 'Preserved.docx') {
  await page.goto('/');
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles({ name, buffer, mimeType: 'application/octet-stream' });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    name.replace('.docx', ''),
  );
}
async function exportDocx(page: Page) {
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download', { timeout: 10000 });
  await page
    .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
    .click();
  const download = await pending;
  return fs.readFile((await download.path())!);
}
test('explicit Word page breaks support keyboard insertion, undo, persistence, export and printed boundaries', async ({
  page,
}) => {
  const source = await fixture();
  await open(page, source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.locator('p').first().click();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Control+ArrowRight');
  await page.keyboard.press('Control+Enter');
  const marker = editor.locator('[data-word-page-break]');
  await expect(marker).toHaveCount(1);
  await expect(marker).toHaveCSS('break-after', 'page');
  await page.keyboard.press('Control+z');
  await expect(marker).toHaveCount(0);
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Page break', exact: true }).click();
  await expect(marker).toHaveCount(1);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Preserved', exact: true }).click();
  await expect(marker).toHaveCount(1);
  const output = await exportDocx(page),
    before = await JSZip.loadAsync(source),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'word/document.xml')
      expect(await after.file(path)!.async('nodebuffer'), path).toEqual(
        await before.file(path)!.async('nodebuffer'),
      );
  expect(await after.file('word/document.xml')!.async('string')).toMatch(
    /<w:br\b[^>]*w:type="page"/,
  );
  await fs.mkdir('.local/word-validation', { recursive: true });
  await fs.writeFile('.local/word-validation/page-break.docx', output);
  await open(page, output, 'Page break result.docx');
  await expect(marker).toHaveCount(1);
  const pdf = await page.pdf({ preferCSSPageSize: true });
  await fs.writeFile('.local/word-validation/page-break.pdf', pdf);
  expect(pdf.toString('latin1').match(/\/Type\s*\/Page\b/g)).toHaveLength(2);
});

test('soft breaks and tabs in an imported Word paragraph survive editing, reload and export', async ({
  page,
}) => {
  const source = await fixture();
  await open(page, source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.click();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+Enter');
  await page.keyboard.type('Inserted line');
  await page.keyboard.press('Tab');
  await page.keyboard.type('After tab');
  await expect(editor.locator('[data-word-tab]')).toHaveCount(1);
  await expect(editor.locator('p').first().locator('br')).toHaveCount(1);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Preserved', exact: true }).click();
  await expect(editor.locator('[data-word-tab]')).toHaveCount(1);
  const output = await exportDocx(page),
    before = await JSZip.loadAsync(source),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'word/document.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const xml = await after.file('word/document.xml')!.async('string');
  expect(xml).toMatch(/<w:br\s*\/>/);
  expect(xml).toMatch(/<w:tab\s*\/>/);
  await open(page, output, 'Breaks result.docx');
  await expect(editor.locator('[data-word-tab]')).toHaveCount(1);
  await expect(editor).toContainText('Inserted line');
});
test('Word paragraph indents and pagination controls persist, undo and preserve imported DOCX parts', async ({
  page,
}) => {
  const source = await fixture();
  await open(page, source);
  const paragraph = page
    .getByRole('textbox', { name: 'Document text', exact: true })
    .locator('p')
    .first();
  await paragraph.click();
  await page.getByRole('button', { name: 'Layout', exact: true }).click();
  for (const [name, value] of [
    ['Indent before text', '36'],
    ['Indent after text', '12'],
    ['First line (negative for hanging)', '-18'],
  ]) {
    await page.getByRole('spinbutton', { name, exact: true }).fill(value);
    await page.getByRole('spinbutton', { name, exact: true }).press('Tab');
  }
  for (const name of ['Keep with next', 'Keep lines together', 'Page break before'])
    await page.getByRole('checkbox', { name, exact: true }).check();
  await page.getByRole('checkbox', { name: 'Widow/orphan control', exact: true }).uncheck();
  await expect(paragraph).toHaveCSS('text-indent', '-24px');
  await expect(paragraph).toHaveCSS('break-before', 'page');
  await paragraph.click();
  await page.keyboard.press('Control+z');
  await expect(
    page.getByRole('checkbox', { name: 'Widow/orphan control', exact: true }),
  ).toBeChecked();
  await page.getByRole('checkbox', { name: 'Widow/orphan control', exact: true }).uncheck();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Preserved', exact: true }).click();
  await expect(paragraph).toHaveCSS('text-indent', '-24px');
  const output = await exportDocx(page),
    after = await JSZip.loadAsync(output),
    before = await JSZip.loadAsync(source);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'word/document.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const xml = await after.file('word/document.xml')!.async('string');
  expect(xml).toContain('w:hanging="360"');
  expect(xml).toContain('w:keepNext w:val="1"');
  expect(xml).toContain('w:widowControl w:val="0"');
  await open(page, output, 'Layout result.docx');
  await expect(paragraph).toHaveCSS('text-indent', '-24px');
  await expect(paragraph).toHaveCSS('break-before', 'page');
});
test('imported Word text and selected formatting edits preserve headers, notes and source page settings', async ({
  page,
}) => {
  const source = await fixture();
  await open(page, source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.locator('p').first().click();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Control+Shift+ArrowRight');
  await page.getByRole('button', { name: 'Underline', exact: true }).click();
  await page.getByRole('button', { name: 'Layout', exact: true }).click();
  await page.getByRole('spinbutton', { name: 'Before paragraph', exact: true }).fill('12');
  await page.getByRole('spinbutton', { name: 'Before paragraph', exact: true }).press('Tab');
  await page.getByRole('spinbutton', { name: 'After paragraph', exact: true }).fill('6');
  await page.getByRole('spinbutton', { name: 'After paragraph', exact: true }).press('Tab');
  await page.getByRole('combobox', { name: 'Text direction', exact: true }).selectOption('rtl');
  await expect(editor.locator('p').first()).toHaveCSS('direction', 'rtl');
  await expect(editor.locator('p').first()).toHaveCSS('margin-top', '16px');
  await editor.locator('p').first().click();
  await page.keyboard.press('Control+Home');
  await page.keyboard.insertText('Edited ');
  await expect(editor.locator('p').first()).toHaveText('Edited Alpha beta');
  const output = await exportDocx(page);
  const original = await JSZip.loadAsync(source),
    changed = await JSZip.loadAsync(output);
  expect(Object.keys(changed.files).sort()).toEqual(Object.keys(original.files).sort());
  for (const [path, part] of Object.entries(original.files))
    if (!part.dir && path !== 'word/document.xml')
      expect(await changed.file(path)!.async('nodebuffer'), path).toEqual(
        await part.async('nodebuffer'),
      );
  const xml = await changed.file('word/document.xml')!.async('string');
  expect(xml).toContain('Edited ');
  expect(xml).toContain('w:u w:val="single"');
  expect(xml).toContain('w:before="240"');
  expect(xml).toContain('w:after="120"');
  expect(xml).toContain('w:bidi w:val="1"');
  expect(xml).toContain('w:left="1320"');
  await open(page, output, 'Roundtrip.docx');
  await expect(editor.locator('p').first()).toHaveText('Edited Alpha beta');
  await expect(editor.locator('p').first()).toHaveCSS('direction', 'rtl');
  await expect(editor.locator('u').first()).toContainText('Alpha');
});

test('footnote text edits update the footnote part and retain the main document and references', async ({
  page,
}) => {
  const source = await fixture();
  await open(page, source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const note = editor
    .locator('p[data-source-paragraph^="1:"]')
    .filter({ hasText: 'Original footnote' });
  await note.click();
  await page.keyboard.press('End');
  await page.keyboard.insertText(' expanded');
  await expect(note).toHaveText('Original footnote expanded');
  const output = await exportDocx(page);
  const original = await JSZip.loadAsync(source),
    changed = await JSZip.loadAsync(output);
  for (const [path, part] of Object.entries(original.files))
    if (!part.dir && path !== 'word/footnotes.xml')
      expect(await changed.file(path)!.async('nodebuffer'), path).toEqual(
        await part.async('nodebuffer'),
      );
  const footnoteText = await page.evaluate((xml) => {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const note = [...doc.getElementsByTagNameNS(ns, 'footnote')]
      .find((element) => element.getAttributeNS(ns, 'id') === '2');
    return note ? [...note.getElementsByTagNameNS(ns, 't')]
      .map((element) => element.textContent).join('') : null;
  }, await changed.file('word/footnotes.xml')!.async('string'));
  expect(footnoteText).toBe('Original footnote expanded');
  await open(page, output, 'EditedNote.docx');
  await expect(editor).toContainText('Original footnote expanded');
});

test('Word document navigation and end additions retain the original body text', async ({
  page,
}) => {
  await open(page, await fixture());
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.click();
  await page.keyboard.press('Control+End');
  await page.keyboard.press('Enter');
  await page.keyboard.insertText('At the end');
  await expect(editor.locator('p').last()).toHaveText('At the end');
  await expect(editor.locator('p').first()).toHaveText('Alpha beta');
  await page.keyboard.press('Control+Shift+Home');
  const selection = await page.evaluate(() => window.getSelection()?.toString());
  expect(selection).toContain('Alpha beta');
  expect(selection).toContain('At the end');
  await page.keyboard.press('ArrowRight');
  const output = await JSZip.loadAsync(await exportDocx(page));
  expect(await output.file('word/document.xml')!.async('string')).toContain('Final body paragraph');
  expect(await output.file('word/footnotes.xml')!.async('string')).toContain('Original footnote');
});
