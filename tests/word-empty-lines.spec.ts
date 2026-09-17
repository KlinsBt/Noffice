import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
test('Word empty hard-break lines use their own fonts through spacing, history, reload and native exports', async ({
  page,
}) => {
  test.setTimeout(150000);
  const root = '.local/word-empty-lines';
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-empty-lines.docx');
  const oracle = JSON.parse(
    await fs.readFile('tests/fixtures/native-word-empty-lines.json', 'utf8'),
  );
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  expect(hash(source)).toBe(oracle.sourceSha256);
  const button = (name: string) => page.getByRole('button', { name, exact: true });
  await page.setViewportSize({ width: 1600, height: 1200 });
  await page.goto('/');
  const upload = async (name: string, buffer: Buffer) => {
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `${name}.docx`,
      buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  await upload('Empty lines', source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const ps = editor.locator('p');
  await expect(ps).toHaveCount(18);
  const heights = () =>
    ps.evaluateAll((nodes) =>
      nodes
        .filter((_, i) => i % 2 === 0)
        .map((p) => {
          const scale = p.getBoundingClientRect().width / parseFloat(getComputedStyle(p).width);
          return (p.getBoundingClientRect().height * 0.75) / scale;
        }),
    );
  const baseline = oracle.stages.find((s: any) => s.stage === 'source').paragraphs;
  const nativeHeights = baseline
    .filter((_: any, i: number) => i % 2 === 0)
    .map((p: any, i: number) => baseline[2 * i + 1].chars[0].y - p.chars[0].y);
  await expect
    .poll(async () => Math.max(...(await heights()).map((v, i) => Math.abs(v - nativeHeights[i]))))
    .toBeLessThan(0.05);
  const lineSizes = [
    [10, 10],
    [10, 10],
    [10, 10, 10],
    [10, 10, 10],
    [10, 10],
    [10, 10],
    [10, 40],
    [10, 10, 10],
    [10, 10, 40],
  ];
  const expected = (name: string) =>
    lineSizes.map((sizes) =>
      sizes.reduce(
        (sum, size) =>
          sum +
          (name === 'minimum'
            ? 50
            : size * 1.15 * (name === 'double' ? 2 : name === 'single' ? 1 : 1.5)),
        0,
      ),
    );
  const check = async (name: string) => {
    await expect
      .poll(async () =>
        Math.max(...(await heights()).map((v, i) => Math.abs(v - expected(name)[i]))),
      )
      .toBeLessThan(0.05);
  };
  const exports: Record<string, string> = {};
  const measured: Record<string, number[]> = { source: await heights() };
  let previous = 'source';
  for (const stage of ['double', 'minimum', 'single']) {
    for (let i = 0; i < 18; i += 2) {
      await ps.nth(i).click();
      await page
        .getByRole('combobox', { name: 'Line spacing', exact: true })
        .selectOption(stage === 'double' ? '2' : stage === 'single' ? '1' : 'custom');
      if (stage === 'minimum') {
        const dialog = page.getByRole('dialog', { name: 'Line spacing', exact: true });
        await dialog
          .getByRole('combobox', { name: 'Spacing rule', exact: true })
          .selectOption('atLeast');
        const amount = dialog.getByRole('spinbutton', { name: 'Spacing amount', exact: true });
        await amount.fill('0');
        await expect(dialog.getByRole('button', { name: 'Apply', exact: true })).toBeDisabled();
        await amount.fill('50');
        await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      }
      await expect
        .poll(async () => (await heights())[i / 2])
        .toBeCloseTo(expected(stage)[i / 2], 1);
      await button('Undo').click();
      await expect
        .poll(async () => (await heights())[i / 2])
        .toBeCloseTo(expected(previous)[i / 2], 1);
      await button('Redo').click();
    }
    await check(stage);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload();
    await button('Recent files').click();
    await button(name).click();
    await check(stage);
    measured[stage] = await heights();
    await button('Export').click();
    const download = page.waitForEvent('download');
    await button('DOCX file Editable in Microsoft Word').click();
    const bytes = await fs.readFile((await (await download).path())!);
    const zip = await JSZip.loadAsync(bytes);
    expect(await zip.file('word/document.xml')!.async('string')).not.toMatch(
      /data-word-|word-line-strut|word-hardline/,
    );
    await fs.writeFile(`${root}/browser-${stage}.docx`, bytes);
    exports[stage] = hash(bytes);
    await upload(`Empty ${stage}`, bytes);
    await check(stage);
    previous = stage;
  }
  await upload('Empty typing', source);
  await check('source');
  await ps.nth(12).click();
  await page.keyboard.press('End');
  await page.keyboard.insertText('New');
  await expect(ps.nth(12)).toContainText('AlphaNew');
  await expect(ps.nth(12).locator('span').filter({ hasText: 'New' }).last()).toHaveCSS(
    'font-size',
    '53.3333px',
  );
  await button('Undo').click();
  await expect(ps.nth(12)).not.toContainText('New');
  await check('source');
  await button('Redo').click();
  await expect(ps.nth(12)).toContainText('New');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await button('Recent files').click();
  await button('Empty typing').click();
  await expect(ps.nth(12)).toContainText('New');
  await button('Export').click();
  const pending = page.waitForEvent('download');
  await button('DOCX file Editable in Microsoft Word').click();
  const typed = await fs.readFile((await (await pending).path())!);
  await fs.writeFile(`${root}/browser-typing.docx`, typed);
  exports.typing = hash(typed);
  await upload('Typed result', typed);
  await expect(ps.nth(12)).toContainText('New');
  // Real caret entry on a leading empty line at half zoom, then restore it.
  for (let i = 0; i < 5; i++) await button('Zoom out').click();
  await ps.nth(0).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Home');
  await page.keyboard.insertText('X');
  await expect(ps.nth(0)).toContainText('X');
  expect(
    await ps.nth(0).evaluate((p) => {
      const range = document.createRange();
      range.setStart(p, 0);
      range.setEndBefore(p.querySelector('br')!);
      return range.toString();
    }),
  ).toBe('X');
  await page.keyboard.press('Control+z');
  await expect(ps.nth(0)).not.toContainText('X');
  // A logical line that wraps must relinquish its one-line measurement.
  await ps.nth(0).click();
  await page.keyboard.press('End');
  await page.keyboard.insertText(' Wide text'.repeat(100));
  await expect(ps.nth(0)).not.toHaveAttribute('data-word-hardlines', 'true');
  await page.keyboard.press('Control+z');
  await expect(ps.nth(0)).toHaveAttribute('data-word-hardlines', 'true');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const stored = await page.evaluate(async () => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('noffice-workspace');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    try {
      return await new Promise<string>((resolve, reject) => {
        const request = db.transaction('files').objectStore('files').getAll();
        request.onsuccess = () => resolve(JSON.stringify(request.result));
        request.onerror = () => reject(request.error);
      });
    } finally {
      db.close();
    }
  });
  expect(stored).toContain('Typed result');
  expect(stored).not.toMatch(/data-word-hardline|data-word-line-strut|--word-measured-leading/);
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify({ passed: true, sourceSha256: hash(source), exports, measured }, null, 2),
  );
});
