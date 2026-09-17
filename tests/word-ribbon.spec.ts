import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';
test('Home case, font stepping, format painter, indents and marks retain imported DOCX content', async ({
  page,
}) => {
  const input = await Packer.toBuffer(
    new Document({
      sections: [
        {
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: 'Source text', bold: true, size: 36, font: 'Georgia' }),
              ],
            }),
            new Paragraph({ children: [new TextRun({ text: 'Target words', italics: true })] }),
          ],
        },
      ],
    }),
  );
  await page.goto('/');
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles({ name: 'Ribbon.docx', mimeType: 'application/octet-stream', buffer: input });
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true }),
    first = editor.locator('p').first(),
    last = editor.locator('p').last();
  await first.click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  await page.getByRole('button', { name: 'Grow font', exact: true }).click();
  await expect(page.getByRole('spinbutton', { name: 'Font size', exact: true })).toHaveValue('20');
  await page.getByRole('button', { name: 'Shrink font', exact: true }).click();
  await expect(page.getByRole('spinbutton', { name: 'Font size', exact: true })).toHaveValue('18');
  await page.getByRole('combobox', { name: 'Change case', exact: true }).selectOption('upper');
  await expect(first).toHaveText('SOURCE TEXT');
  await page.getByRole('button', { name: 'Format painter', exact: true }).click();
  await last.click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  await expect(last.locator('strong')).toHaveText('Target words');
  await expect(last.locator('em')).toHaveCount(0);
  await page.getByRole('combobox', { name: 'Change case', exact: true }).selectOption('lower');
  await expect(last).toHaveText('target words');
  await page.getByRole('button', { name: 'Increase indent', exact: true }).click();
  await expect(last).toHaveCSS('margin-inline-start', '48px');
  await page.keyboard.press('Control+m');
  await expect(last).toHaveCSS('margin-inline-start', '96px');
  await page.keyboard.press('Control+z');
  await expect(last).toHaveCSS('margin-inline-start', '48px');
  await page.getByRole('button', { name: 'Show formatting marks', exact: true }).click();
  await expect(editor.locator('.paragraph-mark')).toHaveCount(2);
  await expect(editor.locator('.word-space-mark')).toHaveCount(2);
  await page.emulateMedia({ media: 'print' });
  await expect(editor.locator('.paragraph-mark').first()).toHaveCSS('display', 'none');
  await expect(editor.locator('.word-space-mark').first()).toHaveCSS('background-image', 'none');
  await page.emulateMedia({ media: 'screen' });
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Ribbon', exact: true }).click();
  await expect(last).toHaveText('target words');
  await expect(last).toHaveCSS('margin-inline-start', '48px');
  await expect(last.locator('strong')).toHaveText('target words');
  await expect(editor.locator('.paragraph-mark')).toHaveCount(0);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!),
    before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'word/document.xml')
      expect(await after.file(path)!.async('nodebuffer'), path).toEqual(
        await before.file(path)!.async('nodebuffer'),
      );
  const xml = await after.file('word/document.xml')!.async('string');
  expect(xml).toContain('SOURCE TEXT');
  expect(xml).toContain('target words');
  expect(xml).toContain('w:start="720"');
  expect(xml).not.toContain('word-format-mark');
  expect(xml).not.toContain('¶');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Ribbon result.docx',
    mimeType: 'application/octet-stream',
    buffer: output,
  });
  await expect(last.locator('strong')).toHaveText('target words');
  await expect(last).toHaveCSS('margin-inline-start', '48px');
});
test('View navigation, ruler and zoom controls and Review word count operate on the current document', async ({
  page,
}) => {
  const input = await Packer.toBuffer(
    new Document({
      sections: [
        {
          children: [
            new Paragraph({ text: 'First heading', heading: 'Heading1' }),
            new Paragraph('Body words here'),
            new Paragraph({ text: 'Second heading', heading: 'Heading2' }),
          ],
        },
      ],
    }),
  );
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'View review.docx',
    mimeType: 'application/octet-stream',
    buffer: input,
  });
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.waitFor();
  await page.getByRole('button', { name: 'View', exact: true }).click();
  await page.getByRole('checkbox', { name: 'Ruler', exact: true }).uncheck();
  await expect(page.locator('.ruler')).toHaveCount(0);
  await page.getByRole('spinbutton', { name: 'Document zoom', exact: true }).fill('120');
  await page.getByRole('spinbutton', { name: 'Document zoom', exact: true }).press('Tab');
  await expect(page.locator('.paper-wrap')).toHaveCSS('zoom', '1.2');
  await page.getByRole('checkbox', { name: 'Navigation pane', exact: true }).check();
  const nav = page.getByRole('navigation', { name: 'Document headings', exact: true });
  await expect(nav.getByRole('button')).toHaveCount(2);
  await nav.getByRole('button', { name: 'Second heading', exact: true }).click();
  await expect(editor).toBeFocused();
  await page.keyboard.press('Shift+End');
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe('Second heading');
  await page.getByRole('button', { name: 'Review', exact: true }).click();
  await page.getByRole('checkbox', { name: 'Browser spellcheck', exact: true }).uncheck();
  await expect(editor).toHaveAttribute('spellcheck', 'false');
  await page.getByRole('checkbox', { name: 'Browser spellcheck', exact: true }).check();
  await page.getByRole('button', { name: 'Word count', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Word count', exact: true });
  await expect(dialog.locator('dd').first()).toHaveText('7');
  await expect(dialog.locator('dd').last()).toHaveText('2');
});
