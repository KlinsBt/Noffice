import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph, TextRun, ExternalHyperlink, Header } from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';

test('Word adds, edits and removes mixed-format links and inserts a cursor link through undo, reload and retained DOCX export', async ({
  page,
}) => {
  const input = await Packer.toBuffer(
    new Document({
      sections: [
        {
          headers: { default: new Header({ children: [new Paragraph('Retained header')] }) },
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: 'Alpha', bold: true }),
                new TextRun({ text: ' beta', italics: true }),
              ],
            }),
            new Paragraph({
              children: [
                new ExternalHyperlink({
                  link: 'https://example.com/original',
                  children: [new TextRun('Source link')],
                }),
              ],
            }),
            new Paragraph('Insert: '),
          ],
        },
      ],
    }),
  );
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Word links.docx',
    mimeType: 'application/octet-stream',
    buffer: input,
  });
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.locator('p').first().click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  expect(await page.evaluate(() => window.getSelection()?.toString())).toBe('Alpha beta');
  await page.keyboard.press('Control+k');
  await expect(page.getByRole('textbox', { name: 'Text to display', exact: true })).toHaveValue(
    'Alpha beta',
  );
  await page
    .getByRole('textbox', { name: 'Link address', exact: true })
    .fill('https://example.com/new');
  await page.getByRole('button', { name: 'Apply link', exact: true }).click();
  const first = editor.locator('p').first();
  await expect(first.locator('a')).toHaveAttribute('href', 'https://example.com/new');
  await expect(first.locator('strong')).toHaveText('Alpha');
  await expect(first.locator('em')).toHaveText(' beta');
  await first.locator('a').click();
  await page.keyboard.press('Control+k');
  await expect(page.getByRole('textbox', { name: 'Text to display', exact: true })).toHaveValue(
    'Alpha beta',
  );
  await page
    .getByRole('textbox', { name: 'Link address', exact: true })
    .fill('https://example.com/revised?a=1&b=2');
  await page.getByRole('button', { name: 'Apply link', exact: true }).click();
  await page.keyboard.press('Control+z');
  await expect(first.locator('a')).toHaveAttribute('href', 'https://example.com/new');
  await page.keyboard.press('Control+y');
  await first.locator('a').click();
  await page.keyboard.press('Control+k');
  await page.getByRole('button', { name: 'Remove link', exact: true }).click();
  await expect(first.locator('a')).toHaveCount(0);
  await page.keyboard.press('Control+z');
  await expect(first.locator('a')).toHaveAttribute('href', 'https://example.com/revised?a=1&b=2');
  await editor.locator('p').last().click();
  await page.keyboard.press('End');
  await expect
    .poll(() =>
      page.evaluate(
        () => window.getSelection()?.anchorNode?.parentElement?.closest('p')?.textContent,
      ),
    )
    .toBe('Insert: ');
  await page.keyboard.press('Control+k');
  await expect(page.getByRole('textbox', { name: 'Text to display', exact: true })).toHaveValue('');
  await page.getByRole('textbox', { name: 'Text to display', exact: true }).fill('Visit now');
  await page
    .getByRole('textbox', { name: 'Link address', exact: true })
    .fill('javascript:alert(1)');
  await page.getByRole('button', { name: 'Apply link', exact: true }).click();
  await expect(page.getByRole('dialog', { name: 'Insert a link', exact: true })).toBeVisible();
  await expect(editor.locator('a')).toHaveCount(2);
  await page
    .getByRole('textbox', { name: 'Link address', exact: true })
    .fill('mailto:team@example.com');
  await page.getByRole('button', { name: 'Apply link', exact: true }).click();
  await page.keyboard.type(' after');
  await expect(editor.locator('p').last().locator('a')).toHaveText('Visit now');
  await expect(editor.locator('p').last()).toContainText('Visit now after');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Word links', exact: true }).click();
  await expect(editor.locator('a')).toHaveCount(3);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.mkdir('.local/docx-validation', { recursive: true });
  await fs.writeFile('.local/docx-validation/links.docx', output);
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (
      !before.files[path].dir &&
      !['word/document.xml', 'word/_rels/document.xml.rels'].includes(path)
    )
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Link result.docx',
    mimeType: 'application/octet-stream',
    buffer: output,
  });
  await expect(editor.locator('a')).toHaveCount(3);
  await expect(first.locator('strong')).toHaveText('Alpha');
  await expect(first.locator('em')).toHaveText(' beta');
  await expect(first.locator('a')).toHaveAttribute('href', 'https://example.com/revised?a=1&b=2');
  await expect(editor.locator('a').nth(1)).toHaveAttribute('href', 'https://example.com/original');
  await expect(editor.locator('a').last()).toHaveAttribute('href', 'mailto:team@example.com');
});
