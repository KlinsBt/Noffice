import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

test('custom Word page geometry renders, presets undo, reload and export with explicit intent', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1600, height: 1200 });
  const root = '.local/word-layout';
  await fs.mkdir(root, { recursive: true });
  const input = await Packer.toBuffer(
    new Document({
      styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
      sections: [
        {
          properties: {
            page: {
              size: { width: 10000, height: 14000 },
              margin: {
                left: 1440,
                right: 1800,
                top: 900,
                bottom: 1000,
                header: 400,
                footer: 500,
                gutter: 0,
              },
            },
          },
          children: [new Paragraph('Custom geometry.')],
        },
      ],
    }),
  );
  await fs.writeFile(`${root}/custom.docx`, input);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Custom geometry.docx',
    buffer: input,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  const title = page.getByRole('textbox', { name: 'File name', exact: true });
  await expect(title).toHaveValue('Custom geometry');
  const paper = page.locator('.paper'),
    editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const geometry = async (width: number, height: number, padding: number[]) => {
    await expect
      .poll(async () => Number.parseFloat(await paper.evaluate((e) => getComputedStyle(e).width)))
      .toBeCloseTo(width / 15, 1);
    await expect
      .poll(async () =>
        Number.parseFloat(await paper.evaluate((e) => getComputedStyle(e).minHeight)),
      )
      .toBeCloseTo(height / 15, 1);
    const actual = await editor.evaluate((e) => {
      const c = getComputedStyle(e);
      return [c.paddingTop, c.paddingRight, c.paddingBottom, c.paddingLeft].map(parseFloat);
    });
    actual.forEach((n, i) => expect(n).toBeCloseTo(padding[i] / 15, 1));
  };
  await geometry(10000, 14000, [900, 1800, 1000, 1440]);
  await page.pdf({ path: `${root}/custom-print.pdf`, preferCSSPageSize: true });
  await page.screenshot({ path: `${root}/custom-screen.png` });
  await page.getByRole('button', { name: 'Layout', exact: true }).click();
  const size = page.getByLabel('Paper size', { exact: true }),
    margin = page.getByLabel('Margins', { exact: true });
  await expect(size).toHaveValue('source');
  await expect(margin).toHaveValue('source');
  // Native controls in word-section-scoped-layout/reference.json bind the
  // A4 preset twips and orientation's asymmetric-margin rotation.
  await size.selectOption('a4');
  await geometry(11907, 16839, [900, 1800, 1000, 1440]);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await geometry(10000, 14000, [900, 1800, 1000, 1440]);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await geometry(11907, 16839, [900, 1800, 1000, 1440]);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.getByLabel('Page orientation', { exact: true }).selectOption('landscape');
  await geometry(14000, 10000, [1440, 900, 1800, 1000]);
  await page.keyboard.press('Control+z');
  await margin.selectOption('normal');
  await geometry(10000, 14000, [1440, 1440, 1440, 1440]);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await geometry(10000, 14000, [900, 1800, 1000, 1440]);
  await size.selectOption('a4');
  await margin.selectOption('normal');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Custom geometry', exact: true }).click();
  await geometry(11907, 16839, [1440, 1440, 1440, 1440]);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.writeFile(`${root}/browser.docx`, output);
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Reimport.docx',
    buffer: output,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  await expect(title).toHaveValue('Reimport');
  await geometry(11907, 16839, [1440, 1440, 1440, 1440]);
  await expect(editor).toContainText('Custom geometry.');
  await page.setViewportSize({ width: 760, height: 1200 });
  await geometry(11907, 16839, [1440, 1440, 1440, 1440]);
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(760);
  expect(
    await page.locator('.document-scroll').evaluate((e) => e.scrollWidth > e.clientWidth),
  ).toBe(true);
  await page.setViewportSize({ width: 1600, height: 1200 });
  await page.pdf({ path: `${root}/browser-print.pdf`, preferCSSPageSize: true });
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        browser: page.context().browser()!.version(),
        sourceHash: createHash('sha256').update(input).digest('hex'),
        exportHash: createHash('sha256').update(output).digest('hex'),
        workflowPassed: true,
      },
      null,
      2,
    ),
  );
});
