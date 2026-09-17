import { test, expect } from '@playwright/test';
import { paginationFixture } from '../scripts/word-pagination-fixture.mjs';
import { wordPlacements } from './word-pagination-helpers';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

test('Word keeps the native caret on an empty continuation page without synthetic editable content', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1500, height: 1200 });
  const root = '.local/word-terminal';
  await fs.mkdir(root, { recursive: true });
  const source = await paginationFixture({ firstLines: 13, terminalBreak: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  await fs.writeFile(`${root}/source.docx`, source);
  await fs.writeFile(`${root}/source.json`, JSON.stringify({ sha256: hash(source) }));
  await page.goto('/');
  const upload = async (name: string, buffer: Buffer) => {
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `${name}.docx`,
      buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  const exportStage = async (name: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const output = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${name}.docx`, output);
    return output;
  };
  await upload('Empty continuation', source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const papers = page.locator('.section-page');
  await expect(papers).toHaveCount(5);
  await expect(editor.locator('p')).toHaveCount(4);
  const positions = await wordPlacements(editor);
  expect(positions[0]).toHaveLength(169);
  expect(positions[0].every((c) => c.page === 1)).toBe(true);
  await editor.locator('p').first().locator('[data-fragment-page="1"]').first().click();
  await page.keyboard.press('Control+Home');
  for (let i = 0; i < 13; i++) await page.keyboard.press('ArrowDown');
  const caret = await page.evaluate(() => {
    const selection = window.getSelection()!;
    const range = selection.getRangeAt(0);
    const collapsedRect = range.getBoundingClientRect();
    // Chromium exposes no collapsed Range rectangle between hard breaks. The
    // real trailing BR supplies the empty line box; insertion below verifies
    // this is the native selection, rather than an overlay caret.
    const next = range.startContainer.childNodes[range.startOffset];
    const box = range.cloneRange();
    if (!collapsedRect.height && next instanceof HTMLBRElement) box.selectNode(next);
    const rect = box.getBoundingClientRect();
    const y = rect.top + rect.height / 2;
    const page =
      [...document.querySelectorAll('.section-page')].findIndex((el) => {
        const r = el.getBoundingClientRect();
        return y >= r.top && y < r.bottom;
      }) + 1;
    return {
      page,
      height: rect.height,
      collapsed: selection.isCollapsed,
      anchor: range.startContainer.nodeName,
      offset: range.startOffset,
      next: next?.nodeName,
      paragraph:
        range.startContainer instanceof Element
          ? range.startContainer.getAttribute('data-source-paragraph')
          : null,
    };
  });
  expect(caret.paragraph).toBe('0:0');
  expect(caret.page).toBe(2);
  expect(caret.height).toBeGreaterThan(0);
  expect(caret.collapsed).toBe(true);
  await page.keyboard.insertText('!');
  await expect
    .poll(async () => (await wordPlacements(editor))[0].at(-1))
    .toEqual({ text: '!', page: 2 });
  const editedPositions = await wordPlacements(editor);
  expect(editedPositions[0].map((c) => c.text).join('')).toBe(
    positions[0].map((c) => c.text).join('') + '!',
  );
  await expect(editor.locator('p')).toHaveCount(4);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(() => wordPlacements(editor)).toEqual(positions);
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('End');
  await page.keyboard.insertText('?');
  await expect
    .poll(async () => (await wordPlacements(editor))[0].find((c) => c.text === '?'))
    .toEqual({ text: '?', page: 1 });
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(() => wordPlacements(editor)).toEqual(positions);
  // Hit the empty line directly, where Word's first paragraph continues on page 2.
  for (let i = 0; i < 5; i++)
    await page.getByRole('button', { name: 'Zoom out', exact: true }).click();
  await expect.poll(() => wordPlacements(editor)).toEqual(positions);
  const second = (await papers.nth(1).boundingBox())!;
  await page.mouse.click(second.x + 14, second.y + 17);
  await page.keyboard.insertText('!');
  await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
  for (let i = 0; i < 5; i++)
    await page.getByRole('button', { name: 'Zoom in', exact: true }).click();
  await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
  const edited = await exportStage('browser-edited');
  await page.pdf({ path: `${root}/browser-edited-print.pdf`, preferCSSPageSize: true });
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(() => wordPlacements(editor)).toEqual(positions);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Empty continuation', exact: true }).click();
  await expect(papers).toHaveCount(5);
  await expect.poll(() => wordPlacements(editor)).toEqual(positions);
  const output = await exportStage('browser');
  expect(hash(output)).toBe(hash(source));
  await page.pdf({ path: `${root}/browser-print.pdf`, preferCSSPageSize: true });
  await page.screenshot({ path: `${root}/screen.png`, fullPage: true });
  await upload('Empty reimport', output);
  await expect(papers).toHaveCount(5);
  await expect.poll(() => wordPlacements(editor)).toEqual(positions);
  await upload('Filled continuation reimport', edited);
  await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceHash: hash(source),
        exportHash: hash(output),
        editedHash: hash(edited),
        browser: page.context().browser()!.version(),
        positions,
        editedPositions,
        caret,
      },
      null,
      2,
    ),
  );
});
