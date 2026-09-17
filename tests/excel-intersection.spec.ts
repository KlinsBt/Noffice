import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
const native: typeof import('./fixtures/native-excel-intersection.json') = JSON.parse(
  (await fs.readFile('tests/fixtures/native-excel-intersection.json', 'utf8')).replace(
    /^\uFEFF/,
    '',
  ),
);

test('scalar intersections retain blanks, errors and caller coordinates through history, reload and exports', async ({
  page,
}) => {
  test.setTimeout(90000);
  const root = '.local/excel-intersection';
  const source = await fs.readFile('tests/fixtures/excel-intersection.xlsx');
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  expect(hash(source)).toBe(native.sourceSha256);
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
      ref: string;
      value: number | string | boolean | null;
      kind: string;
    }[],
  ) => {
    for (const sheet of ['Calls']) {
      await button(sheet).click();
      for (const item of cells) {
        const value =
          typeof item.value === 'boolean'
            ? String(item.value).toUpperCase()
            : String(item.value ?? '');
        await expect(cell(item.ref)).toHaveText(value);
      }
    }
  };
  const exports: Record<string, string> = {};
  await page.goto('/');
  for (const stage of native.stages) {
    const name = `Intersection ${stage.stage}`;
    await upload(source, name);
    await check(native.initial);
    await button('Data').click();
    const target = 'A2';
    await cell(target).dblclick();
    await page
      .getByRole('textbox', { name: `Edit ${target}`, exact: true })
      .fill(stage.stage === 'error' ? '=1/0' : stage.stage === 'edit' ? '64' : '');
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
    await fs.writeFile(`${root}/${stage.stage}.xlsx`, output);
    await upload(output, `Reimport ${stage.stage}`);
    await check(stage.cells);
    exports[stage.stage] = hash(output);
  }
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceSha256: hash(source),
        exports,
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ),
  );
});
