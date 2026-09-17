import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph, Header, Footer } from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

test('Word tracks live section splits and joins through undo, reload and native export', async ({
  page,
}) => {
  const zip = await JSZip.loadAsync(
    await Packer.toBuffer(
      new Document({
        evenAndOddHeaderAndFooters: true,
        sections: [
          {
            headers: {
              default: new Header({ children: [new Paragraph('Odd header')] }),
              even: new Header({ children: [new Paragraph('Even header')] }),
              first: new Header({ children: [new Paragraph('First header')] }),
            },
            footers: {
              default: new Footer({ children: [new Paragraph('Odd footer')] }),
              even: new Footer({ children: [new Paragraph('Even footer')] }),
              first: new Footer({ children: [new Paragraph('First footer')] }),
            },
            children: [new Paragraph('Fixture')],
          },
        ],
      }),
    ),
  );
  const references = (
    (await zip.file('word/document.xml')!.async('string')).match(
      /<w:(?:headerReference|footerReference)\b[^>]*\/>/g,
    ) || []
  ).join('');
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:pPr><w:sectPr>${references}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>AlphaBeta</w:t></w:r></w:p><w:p><w:r><w:t>Gamma</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr></w:body></w:document>`,
  );
  const input = await zip.generateAsync({ type: 'nodebuffer' });
  const root = '.local/word-ranges';
  await fs.mkdir(root, { recursive: true });
  await fs.writeFile(`${root}/browser-source.docx`, input);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Section ranges.docx',
    buffer: input,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const status = page.getByLabel('Selected sections');
  const exportStage = async (name: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const output = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${name}.docx`, output);
    return createHash('sha256').update(output).digest('hex');
  };
  await editor.locator('p').first().click();
  await expect(status).toHaveText('Section 1 of 2');
  await editor.locator('p').nth(1).click();
  await expect(status).toHaveText('Section 2 of 2');
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Control+Shift+End');
  await expect(status).toHaveText('Sections 1–2 of 2');
  await page.getByRole('button', { name: 'Show formatting marks', exact: true }).click();
  await expect(editor.locator('.section-mark')).toHaveCount(1);
  await expect(editor.locator('.section-mark')).toHaveAttribute('data-symbol', 'Section break');
  await page.screenshot({ path: `${root}/section-marks.png` });
  await editor.locator('p').nth(1).click();
  await page.keyboard.press('End');
  await page.keyboard.type(' edited');
  await expect(status).toHaveText('Section 2 of 2');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Section ranges', exact: true }).click();
  await editor.locator('p').nth(1).click();
  await expect(status).toHaveText('Section 2 of 2');
  await expect(editor.locator('p').nth(1)).toHaveText('Gamma edited');
  await page.getByRole('button', { name: 'Show formatting marks', exact: true }).click();
  await editor.locator('p').first().click();
  await page.keyboard.press('Home');
  for (let i = 0; i < 5; i++) await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Enter');
  await expect(editor.locator('p').first()).toHaveText('Alpha');
  await expect(editor.locator('p').nth(1)).toHaveText('Beta');
  await expect(status).toHaveText('Section 1 of 2');
  await expect(editor.locator('.section-mark')).toHaveCount(1);
  const splitHash = await exportStage('browser-split');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const reopened = await page.context().newPage();
  await reopened.goto('/');
  await reopened.getByRole('button', { name: 'Recent files', exact: true }).click();
  await reopened.getByRole('button', { name: 'Section ranges', exact: true }).click();
  await reopened
    .getByRole('textbox', { name: 'Document text', exact: true })
    .locator('p')
    .nth(1)
    .click();
  await expect(reopened.getByLabel('Selected sections')).toHaveText('Section 1 of 2');
  await reopened.close();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor.locator('p').first()).toHaveText('AlphaBeta');
  await expect(editor.locator('.section-mark')).toHaveCount(1);
  await expect(status).toHaveText('Section 1 of 2');
  await editor.locator('p').nth(1).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Backspace');
  await expect(editor.locator('p').first()).toHaveText('AlphaBetaGamma edited');
  await expect(status).toHaveText('Section 1 of 1');
  await expect(editor.locator('.section-mark')).toHaveCount(0);
  const joinHash = await exportStage('browser-join');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor.locator('p').first()).toHaveText('AlphaBeta');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const download = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
    .click();
  const output = await fs.readFile((await (await download).path())!);
  await fs.writeFile(`${root}/browser.docx`, output);
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Section ranges exported.docx',
    buffer: output,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Section ranges exported',
  );
  await expect(editor.locator('p').nth(1)).toHaveText('Gamma edited');
  await editor.locator('p').first().click();
  await expect(status).toHaveText('Section 1 of 2');
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        workflowPassed: true,
        sourceHash: createHash('sha256').update(input).digest('hex'),
        exportHash: createHash('sha256').update(output).digest('hex'),
        splitHash,
        joinHash,
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ),
  );
});
