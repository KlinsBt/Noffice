import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
const native: typeof import('./fixtures/native-excel-sumproduct.json') = JSON.parse(
  (await fs.readFile('tests/fixtures/native-excel-sumproduct.json', 'utf8')).replace(/^\uFEFF/, ''),
);

test('SUMPRODUCT retains matrix shapes and typed arithmetic through history, reload and exports', async ({
  page,
}) => {
  test.setTimeout(150000);
  const root = '.local/excel-sumproduct';
  const source = await fs.readFile('tests/fixtures/excel-sumproduct.xlsx');
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
      const expected = cells.map((item) => ({
        ref: item.ref,
        count: 1,
        value:
          typeof item.value === 'boolean'
            ? String(item.value).toUpperCase()
            : typeof item.value === 'number'
              ? String(Number(item.value.toPrecision(12)))
              : String(item.value ?? ''),
      }));
      // Keep every independently captured native value and strict cell identity.
      // One snapshot avoids thousands of serial protocol/trace round trips.
      await expect
        .poll(
          () =>
            page.getByRole('gridcell').evaluateAll(
              (nodes, refs) => {
                const values = new Map<string, string[]>();
                for (const node of nodes) {
                  const ref = node.getAttribute('aria-label') || '';
                  const entries = values.get(ref) || [];
                  entries.push(node.textContent || '');
                  values.set(ref, entries);
                }
                return refs.map((ref) => ({
                  ref,
                  count: values.get(ref)?.length || 0,
                  value: values.get(ref)?.[0] ?? null,
                }));
              },
              cells.map((item) => item.ref),
            ),
          { timeout: 5000 },
        )
        .toEqual(expected);
    }
  };
  const exports: Record<string, string> = {};
  await page.goto('/');
  for (const stage of native.stages) {
    const name = `Sumproduct ${stage.stage}`;
    await upload(source, name);
    await check(native.initial);
    if (stage.stage === 'edit') {
      await page.screenshot({ path: `${root}/browser-source.png` });
      await cell('A1').dblclick();
      await page.getByRole('textbox', { name: 'Edit A1', exact: true }).fill('=SUMPRODUCT()');
      await page.keyboard.press('Enter');
      await expect(cell('A1')).toHaveText('#VALUE!');
      await button('Undo').click();
      await check(native.initial);
    }
    await button(stage.stage === 'formula' ? 'Calls' : 'Data').click();
    const target = stage.stage === 'formula' ? 'D8' : stage.stage === 'table' ? 'H3' : 'A2';
    await cell(target).dblclick();
    await page
      .getByRole('textbox', { name: `Edit ${target}`, exact: true })
      .fill(
        stage.stage === 'formula'
          ? '=SUMPRODUCT((Data!A1:A3>=4)*Data!B1:B3)'
          : stage.stage === 'table'
            ? '4'
            : stage.stage === 'error'
              ? '=1/0'
              : stage.stage === 'edit'
                ? '64'
                : '',
      );
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
