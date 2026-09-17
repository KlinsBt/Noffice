import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

test('uniform Word text places automatic leading below and minimum leading above the text', async ({
  page,
}) => {
  const root = '.local/word-uniform-leading';
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-baselines.docx');
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  const oracle = JSON.parse(
    (await fs.readFile('tests/fixtures/native-word-uniform-leading.json', 'utf8')).replace(
      /^\uFEFF/,
      '',
    ),
  );
  expect(hash(source)).toBe(oracle.sourceSha256);
  await fs.writeFile(`${root}/source.docx`, source);
  await page.setViewportSize({ width: 1500, height: 1200 });
  await page.goto('/');
  const upload = async (name: string, buffer: Buffer) => {
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `${name}.docx`,
      buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  await upload('Uniform leading', source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const paragraphs = editor.locator('p');
  await expect(paragraphs).toHaveCount(18);
  const target = paragraphs.nth(9);
  const measure = () =>
    target.evaluate((p) => {
      const text = document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode()!;
      const range = document.createRange();
      range.setStart(text, 0);
      range.setEnd(text, 1);
      const rect = p.getBoundingClientRect();
      return {
        height: rect.height * 0.75,
        textTop: (range.getBoundingClientRect().top - rect.top) * 0.75,
      };
    });
  await expect.poll(async () => (await measure()).height).toBeCloseTo(17.25, 1);
  const initial = await measure();
  const sourceHeights = await paragraphs.evaluateAll((ps) =>
    ps.slice(9, 12).map((p) => p.getBoundingClientRect().height),
  );
  const stages: Record<string, { height: number; textTop: number }> = {};
  const exports: Record<string, string> = {};
  const check = async (height: number, delta: number) => {
    await expect.poll(async () => (await measure()).height).toBeCloseTo(height, 1);
    // Native same-font text origins remain fixed for automatic multiples; minimum
    // spacing puts its excess above the text. Absolute glyph baselines are separate.
    await expect
      .poll(async () => Math.abs((await measure()).textTop - initial.textTop - delta))
      .toBeLessThan(0.15);
  };
  const save = async (stage: string, height: number, delta: number) => {
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    await check(height, delta);
    stages[stage] = await measure();
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/browser-${stage}.docx`, bytes);
    exports[stage] = hash(bytes);
    await upload(`Uniform ${stage}`, bytes);
    await check(height, delta);
  };
  await target.click();
  await page.getByRole('combobox', { name: 'Line spacing', exact: true }).selectOption('2');
  await check(23, 0);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await check(17.25, 0);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await check(23, 0);
  await save('double', 23, 0);
  await target.click();
  await page.getByRole('combobox', { name: 'Line spacing', exact: true }).selectOption('custom');
  const dialog = page.getByRole('dialog', { name: 'Line spacing', exact: true });
  await dialog.getByRole('combobox', { name: 'Spacing rule', exact: true }).selectOption('atLeast');
  const amount = dialog.getByRole('spinbutton', { name: 'Spacing amount', exact: true });
  await amount.fill('0');
  await expect(dialog.getByRole('button', { name: 'Apply', exact: true })).toBeDisabled();
  await check(23, 0);
  await amount.fill('50');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await check(50, 38.5);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await check(23, 0);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await save('minimum', 50, 38.5);
  await target.click();
  await page.getByRole('combobox', { name: 'Line spacing', exact: true }).selectOption('1');
  await check(11.5, 0);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await check(50, 38.5);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await save('single', 11.5, 0);
  // Uniform wrapped text retains the earlier homogeneous advance path and
  // recovers the single-line placement when enough width returns.
  await editor.evaluate((el) => (el.style.width = '30px'));
  await expect(target).not.toHaveAttribute('data-word-mixed-leading');
  await expect.poll(async () => (await measure()).height).toBeGreaterThan(11.5);
  await editor.evaluate((el) => el.style.removeProperty('width'));
  await expect(target).toHaveAttribute('data-word-mixed-leading', 'true');
  await check(11.5, 0);
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceSha256: hash(source),
        initial,
        sourceHeights,
        stages,
        exports,
      },
      null,
      2,
    ),
  );
});
