import { wordPlacements } from './word-pagination-helpers';
import { test, expect } from '@playwright/test';
import { paginationFixture } from '../scripts/word-pagination-fixture.mjs';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

test('Word paginates intact paragraphs, keeps lines together, edits across pages and exports without fragment markup', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1500, height: 1200 });
  const root = '.local/word-pagination';
  await fs.mkdir(root, { recursive: true });
  const source = await paginationFixture();
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
  await upload('Pagination', source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true }),
    papers = page.locator('.section-page');
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
  await expect(papers).toHaveCount(6);
  await expect(editor.locator('p')).toHaveCount(4);
  // Native Word: 13/7 lines, kept paragraph starts on page 3, wrapping uses pages 3-5.
  const placement = () => wordPlacements(editor);
  const positions = await placement();
  expect(positions[0].slice(0, 169).every((c) => c.page === 1)).toBe(true);
  expect(positions[0].slice(169).every((c) => c.page === 2)).toBe(true);
  expect(positions[1].every((c) => c.page === 3)).toBe(true);
  expect(positions[2].slice(0, 200).every((c) => c.page === 3)).toBe(true);
  expect(positions[2].slice(200, 720).every((c) => c.page === 4)).toBe(true);
  expect(positions[2].slice(720).every((c) => c.page === 5)).toBe(true);
  expect(positions[3].every((c) => c.page === 6)).toBe(true);
  await editor.click();
  await page.keyboard.press('Control+Home');
  for (let i = 0; i < 13; i++) await page.keyboard.press('ArrowDown');
  await page.keyboard.insertText('@');
  expect((await placement())[0].find((c) => c.text === '@')?.page).toBe(2);
  await page.keyboard.press('Control+z');
  await expect(papers).toHaveCount(6);
  // A real pointer hit on a later page must focus the original paragraph.
  const later = editor
    .locator('p')
    .first()
    .locator('[data-fragment-page="2"]')
    .filter({ hasText: 'Line 14' })
    .first();
  await later.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' edit');
  await expect(editor.locator('p').first()).toContainText(' edit');
  expect(
    (await placement())[0].filter((c) => c.text === 'e' && c.page === 2).length,
  ).toBeGreaterThan(0);
  await expect(editor.locator('p')).toHaveCount(4);
  await page.keyboard.press('Control+z');
  await expect(editor.locator('p').first()).not.toContainText(' edit');
  await expect(papers).toHaveCount(6);
  await editor.locator('p').last().click();
  await page.keyboard.press('End');
  await page.keyboard.type('!');
  const edited = await exportStage('browser-edited');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Pagination', exact: true }).click();
  await expect(papers).toHaveCount(6);
  expect(await placement()).toEqual(positions);
  const output = await exportStage('browser');
  await page.pdf({ path: `${root}/browser-print.pdf`, preferCSSPageSize: true });
  await page.screenshot({ path: `${root}/screen.png` });
  await upload('Pagination reimport', output);
  await expect(papers).toHaveCount(6);
  await expect(editor.locator('p')).toHaveCount(4);
  expect(await placement()).toEqual(positions);
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
      },
      null,
      2,
    ),
  );
});
