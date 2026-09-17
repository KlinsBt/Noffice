import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

test('large paragraph marks stay outside nonempty line metrics and return for empty typing', async ({
  page,
}) => {
  const root = '.local/word-baselines';
  const source = await fs.readFile('tests/fixtures/word-baselines.docx');
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  const oracle = JSON.parse(
    (await fs.readFile('tests/fixtures/native-word-baselines.json', 'utf8')).replace(/^\uFEFF/, ''),
  );
  expect(hash(source)).toBe(oracle.sourceSha256);
  await page.setViewportSize({ width: 1500, height: 1200 });
  await page.goto('/');
  const upload = async (name: string, buffer: Buffer) => {
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({
        name: `${name}.docx`,
        buffer,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  await upload('Paragraph marks', source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const paragraphs = editor.locator('p');
  await expect(paragraphs).toHaveCount(18);
  const target = paragraphs.nth(12);
  const sourceHeights: number[] = [];
  for (let i = 12; i < 15; i++) {
    const advance = oracle.source[i + 1].chars[0].y - oracle.source[i].chars[0].y;
    await expect
      .poll(async () => Math.abs((await paragraphs.nth(i).boundingBox())!.height * 0.75 - advance))
      .toBeLessThan(0.15);
    sourceHeights.push((await paragraphs.nth(i).boundingBox())!.height);
  }
  await target.click();
  await page.keyboard.press('Home');
  for (let i = 0; i < 3; i++) await page.keyboard.press('ArrowRight');
  for (let i = 0; i < 5; i++) await page.keyboard.press('Shift+ArrowRight');
  expect(await page.evaluate(() => getSelection()?.toString())).toBe('Large');
  const size = page.getByRole('spinbutton', { name: 'Font size', exact: true });
  await size.fill('20');
  await size.press('Enter');
  const height = async (expected: number) =>
    expect.poll(async () => (await target.boundingBox())!.height * 0.75).toBeCloseTo(expected, 1);
  await height(34.5);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await height(17.25);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await height(34.5);
  const exports: Record<string, string> = {};
  const heights: Record<string, number> = {};
  const save = async (stage: string, text: string, expected: number) => {
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    await expect(target).toHaveText(text);
    await height(expected);
    heights[stage] = (await target.boundingBox())!.height * 0.75;
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/browser-${stage}.docx`, bytes);
    exports[stage] = hash(bytes);
    await upload(`Marks ${stage}`, bytes);
    await expect(target).toHaveText(text);
    await height(expected);
  };
  await save('size', 'Sg Large', 34.5);
  await target.click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  await page.keyboard.press('Backspace');
  await expect(target).toHaveText('');
  await height(69);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await height(34.5);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await height(69);
  await save('empty', '', 69);
  await target.click();
  await expect(size).toHaveValue('40');
  await page.keyboard.insertText('New');
  await expect(target).toHaveText('New');
  await height(69);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(target).toHaveText('');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(target).toHaveText('New');
  await save('typing', 'New', 69);
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify({ passed: true, sourceSha256: hash(source), exports, heights, sourceHeights }, null, 2),
  );
});
