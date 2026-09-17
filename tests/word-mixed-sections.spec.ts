import { test, expect } from '@playwright/test';
import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

test('mixed Word sections keep distinct editable page geometry through split, join, undo, reload and print', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1600, height: 1200 });
  const root = '.local/word-mixed';
  await fs.mkdir(root, { recursive: true });
  const zip = await JSZip.loadAsync(
    await Packer.toBuffer(
      new Document({
        styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
        sections: [{ children: [new Paragraph('Fixture')] }],
      }),
    ),
  );
  const first =
    '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="900" w:right="1800" w:bottom="1000" w:left="1440" w:header="400" w:footer="400" w:gutter="0"/>';
  const last =
    '<w:type w:val="nextPage"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1080" w:right="900" w:bottom="720" w:left="720" w:header="400" w:footer="400" w:gutter="0"/>';
  const spacing = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr>${spacing}<w:sectPr>${first}</w:sectPr></w:pPr><w:r><w:t>AlphaBeta</w:t></w:r></w:p><w:p><w:pPr>${spacing}</w:pPr><w:r><w:t>Gamma</w:t></w:r></w:p><w:sectPr>${last}</w:sectPr></w:body></w:document>`,
  );
  const input = await zip.generateAsync({ type: 'nodebuffer' });
  await fs.writeFile(`${root}/source.docx`, input);
  await page.goto('/');
  const upload = async (name: string, buffer: Buffer) => {
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `${name}.docx`,
      buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  await upload('Mixed sections', input);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true }),
    papers = page.locator('.section-page');
  const geometry = async () => {
    await expect(papers).toHaveCount(2);
    const rects = await papers.evaluateAll((es) =>
      es.map((e) => {
        const r = e.getBoundingClientRect();
        return { x: r.x, y: r.y, width: r.width, height: r.height };
      }),
    );
    for (const [i, w, h] of [
      [0, 12240, 15840],
      [1, 16838, 11906],
    ]) {
      expect(rects[i].width).toBeCloseTo(w / 15, 1);
      expect(rects[i].height).toBeCloseTo(h / 15, 1);
    }
    expect(rects[1].y - rects[0].y).toBeCloseTo(1056 + 24, 1);
    const blocks = await editor.locator('p').evaluateAll((es) =>
      es.map((e) => {
        const r = e.getBoundingClientRect();
        return { x: r.x, y: r.y, section: Number(e.getAttribute('data-surface-section')) };
      }),
    );
    expect(blocks[0].x - rects[0].x).toBeCloseTo(96, 1);
    expect(blocks[0].y - rects[0].y).toBeCloseTo(60, 1);
    const second = blocks.find((b) => b.section === 1)!;
    expect(second.x - rects[1].x).toBeCloseTo(48, 1);
    expect(second.y - rects[1].y).toBeCloseTo(72, 1);
    return { rects, blocks };
  };
  const initial = await geometry();
  for (let i = 0; i < 5; i++)
    await page.getByRole('button', { name: 'Zoom out', exact: true }).click();
  await expect.poll(async () => (await papers.first().boundingBox())!.width).toBeCloseTo(408, 1);
  await page.screenshot({ path: `${root}/screen.png`, fullPage: true });
  for (let i = 0; i < 5; i++)
    await page.getByRole('button', { name: 'Zoom in', exact: true }).click();
  await geometry();
  await page.setViewportSize({ width: 680, height: 1200 });
  await expect(papers).toHaveCount(0);
  await expect(editor).toContainText('AlphaBeta');
  await page.setViewportSize({ width: 1600, height: 1200 });
  await geometry();
  // Real pointer hit testing must resolve to the second section despite absolute page positions.
  await editor.locator('p').nth(1).click();
  await expect(page.getByLabel('Selected sections')).toHaveText('Section 2 of 2');
  await page.keyboard.press('End');
  await page.keyboard.type(' edited');
  await editor.locator('p').first().click();
  await page.keyboard.press('Home');
  for (let i = 0; i < 5; i++) await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Enter');
  await expect(editor.locator('p')).toHaveCount(3);
  await expect(editor.locator('p').nth(1)).toHaveText('Beta');
  await geometry();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor.locator('p').first()).toHaveText('AlphaBeta');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(editor.locator('p').nth(1)).toHaveText('Beta');
  await geometry();
  const exportStage = async (name: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${name}.docx`, bytes);
    return bytes;
  };
  const split = await exportStage('browser-split');
  await page.pdf({ path: `${root}/split-print.pdf`, preferCSSPageSize: true });
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Mixed sections', exact: true }).click();
  await geometry();
  await expect(editor.locator('p').nth(1)).toHaveText('Beta');
  await editor.locator('p').nth(2).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Backspace');
  await expect(editor.locator('p').nth(1)).toHaveText('BetaGamma edited');
  await expect(papers).toHaveCount(0);
  await expect(page.locator('.paper-wrap')).toHaveClass(/source-layout/);
  expect(
    await page.locator('.paper').evaluate((e) => parseFloat(getComputedStyle(e).width)),
  ).toBeCloseTo(16838 / 15, 1);
  const joined = await exportStage('browser-join');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await geometry();
  // No decorations enter the source document; a real DOCX reimport reconstructs geometry.
  await upload('Mixed reimport', split);
  await geometry();
  await page.setViewportSize({ width: 760, height: 1200 });
  await geometry();
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(760);
  expect(
    await page.locator('.document-scroll').evaluate((e) => e.scrollWidth > e.clientWidth),
  ).toBe(true);
  await page.setViewportSize({ width: 1600, height: 1200 });
  // Unsupported long flow must not masquerade as one page per section; Undo recovers layout.
  await editor.locator('p').nth(2).click();
  await page.keyboard.press('End');
  await page.keyboard.insertText(' Long wrapping text.'.repeat(700));
  await expect(papers).toHaveCount(0);
  await expect(
    page.getByText('Continuous view: this content needs paragraph pagination.', { exact: true }),
  ).toBeVisible();
  await page.keyboard.press('Control+z');
  await geometry();
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        workflowPassed: true,
        browser: page.context().browser()!.version(),
        sourceHash: hash(input),
        splitHash: hash(split),
        joinHash: hash(joined),
        printHash: hash(await fs.readFile(`${root}/split-print.pdf`)),
        initial,
      },
      null,
      2,
    ),
  );
});
