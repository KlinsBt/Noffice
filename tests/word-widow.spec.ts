import { test, expect } from '@playwright/test';
import { paginationFixture } from '../scripts/word-pagination-fixture.mjs';
import { wordPlacements } from './word-pagination-helpers';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

for (const orphan of [false, true])
  test(`Word ${orphan ? 'orphan' : 'widow'} control repaginates the original paragraph with atomic history and retained export`, async ({
    page,
  }) => {
    await page.setViewportSize({ width: 1500, height: 1200 });
    const root = orphan ? '.local/word-orphan' : '.local/word-widow';
    const target = orphan ? 1 : 0,
      beforeBreak = orphan ? 13 : 169,
      afterBreak = orphan ? 0 : 156;
    await fs.mkdir(root, { recursive: true });
    const source = await paginationFixture(
      orphan ? { firstLines: 12, keepSecond: false } : { firstLines: 14 },
    );
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
    await upload('Widow control', source);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    const papers = page.locator('.section-page');
    await expect(papers).toHaveCount(5);
    await expect(editor.locator('p')).toHaveCount(4);
    const positions = await wordPlacements(editor);
    expect(positions[target].slice(0, beforeBreak).every((c) => c.page === 1)).toBe(true);
    expect(positions[target].slice(beforeBreak).every((c) => c.page === 2)).toBe(true);
    await editor.locator('p').nth(target).locator('[data-fragment-page]').first().click();
    await page.keyboard.press('Home');
    await page.getByRole('button', { name: 'Layout', exact: true }).click();
    const widow = page.getByRole('checkbox', { name: 'Widow/orphan control', exact: true });
    await expect(widow).not.toBeChecked();
    await widow.check();
    const assertEnabled = async () => {
      await expect
        .poll(async () => (await wordPlacements(editor))[target][afterBreak].page)
        .toBe(2);
      await expect(papers).toHaveCount(5);
      await expect(editor.locator('p')).toHaveCount(4);
      const actual = await wordPlacements(editor);
      expect(actual[target].slice(0, afterBreak).every((c) => c.page === 1)).toBe(true);
      expect(actual[target].slice(afterBreak).every((c) => c.page === 2)).toBe(true);
      expect(actual.map((p) => p.map((c) => c.text).join(''))).toEqual(
        positions.map((p) => p.map((c) => c.text).join('')),
      );
      return actual;
    };
    const editedPositions = await assertEnabled();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(widow).not.toBeChecked();
    await expect.poll(() => wordPlacements(editor)).toEqual(positions);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(widow).toBeChecked();
    await assertEnabled();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: 'Widow control', exact: true }).click();
    await assertEnabled();
    const edited = await exportStage('browser-edited');
    await page.pdf({ path: `${root}/browser-edited-print.pdf`, preferCSSPageSize: true });
    await upload('Widow enabled reimport', edited);
    expect(await assertEnabled()).toEqual(editedPositions);
    await editor.locator('p').nth(target).locator('[data-fragment-page]').first().click();
    await page.keyboard.press('Home');
    await page.getByRole('button', { name: 'Layout', exact: true }).click();
    await expect(widow).toBeChecked();
    await widow.uncheck();
    await expect.poll(() => wordPlacements(editor)).toEqual(positions);
    const output = await exportStage('browser');
    await page.pdf({ path: `${root}/browser-print.pdf`, preferCSSPageSize: true });
    await upload('Widow disabled reimport', output);
    await expect.poll(() => wordPlacements(editor)).toEqual(positions);
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
        },
        null,
        2,
      ),
    );
  });
