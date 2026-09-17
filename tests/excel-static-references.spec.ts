import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
const native: typeof import('./fixtures/native-excel-static-references.json') = JSON.parse(
  (await fs.readFile('tests/fixtures/native-excel-static-references.json', 'utf8')).replace(
    /^\uFEFF/,
    '',
  ),
);

test('static references and blank results survive edits, history, reload and three actual exports', async ({
  page,
}) => {
  test.setTimeout(90000);
  const root = '.local/excel-static-references';
  const source = await fs.readFile('tests/fixtures/excel-static-references.xlsx');
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  expect(hash(source)).toBe(native.sourceHash);
  await fs.mkdir(root, { recursive: true });
  await fs.writeFile(`${root}/source.xlsx`, source);
  const button = (name: string) => page.getByRole('button', { name, exact: true });
  const cell = (name: string) => page.getByRole('gridcell', { name, exact: true });
  const upload = async (buffer: Buffer, name: string) => {
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({ name: `${name}.xlsx`, buffer, mimeType: 'application/octet-stream' });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  const check = async (
    cells: {
      sheet: string;
      address: string;
      value: number | string | boolean | null;
      valueType: string;
    }[],
  ) => {
    for (const sheet of ['Data', 'Cost data', 'Summary']) {
      await button(sheet).click();
      for (const item of cells.filter((c) => c.sheet === sheet)) {
        const value =
          item.valueType === 'System.Int32'
            ? '#DIV/0!'
            : typeof item.value === 'boolean'
              ? String(item.value).toUpperCase()
              : String(item.value ?? '');
        await expect(cell(item.address)).toHaveText(value);
      }
    }
  };
  const exports: { stage: string; sha256: string }[] = [];
  await page.goto('/');
  for (const stage of native.stages) {
    const name = `Static ${stage.stage}`;
    await upload(source, name);
    await check(native.initial);
    await button('Data').click();
    const target = stage.stage === 'branch' ? 'B2' : 'A1';
    await cell(target).dblclick();
    await page
      .getByRole('textbox', { name: `Edit ${target}`, exact: true })
      .fill(stage.stage === 'branch' ? '9' : stage.stage === 'edit' ? '5' : '');
    await page.keyboard.press('Enter');
    await check(stage.cells);
    await button('Undo').click();
    await check(native.initial);
    await button('Redo').click();
    await check(stage.cells);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await button('Recent files').click();
    await button(name).click();
    await check(stage.cells);
    await button('Export').click();
    const pending = page.waitForEvent('download');
    await button('XLSX file Editable in Microsoft Excel').click();
    const output = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/browser-${stage.stage}.xlsx`, output);
    await upload(output, `Reimport ${stage.stage}`);
    await check(stage.cells);
    exports.push({ stage: stage.stage, sha256: hash(output) });
  }
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceHash: hash(source),
        exports,
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ),
  );
});
