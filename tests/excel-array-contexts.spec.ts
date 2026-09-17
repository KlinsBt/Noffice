import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
const native = JSON.parse(
  await fs.readFile('tests/fixtures/native-excel-array-contexts.json', 'utf8'),
);
test('CSE entry identity survives real edits, history, failures, reload and actual exports', async ({
  page,
}) => {
  test.setTimeout(180000);
  const root = '.local/excel-arrays',
    source = await fs.readFile('tests/fixtures/excel-array-contexts.xlsx');
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  expect(hash(source)).toBe(native.sourceSha256);
  await fs.mkdir(root, { recursive: true });
  const button = (name: string) => page.getByRole('button', { name, exact: true });
  const cell = (name: string) => page.getByRole('gridcell', { name, exact: true });
  const upload = async (buffer: Buffer, name: string) => {
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({ name: `${name}.xlsx`, buffer, mimeType: 'application/octet-stream' });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  const check = async (rows: any[]) => {
    await button('Calls').click();
    for (const row of rows) await expect(cell(row.ref)).toHaveText(String(row.value));
  };
  const exports: Record<string, string> = {};
  await page.setViewportSize({ width: 1500, height: 1100 });
  await page.goto('/');
  for (const stage of native.editingStages) {
    const name = `Arrays ${stage.stage}`;
    await upload(source, name);
    await check(native.stages[0].cells);
    if (stage.stage === 'edit') {
      await cell('B1').dblclick();
      await page.getByRole('textbox', { name: 'Edit B1', exact: true }).fill('text');
      await page.keyboard.press('Control+Shift+Enter');
      await expect(
        page.getByText('Enter a formula before using Ctrl+Shift+Enter.', { exact: true }),
      ).toBeVisible();
      await check(native.stages[0].cells);
      await cell('B1').click();
      await cell('C1').click({ modifiers: ['Shift'] });
      await page.getByRole('textbox', { name: 'Formula bar', exact: true }).fill('=3');
      await page.keyboard.press('Control+Shift+Enter');
      await expect(
        page.getByText('Multi-cell array entry is not supported yet. Select one cell.', {
          exact: true,
        }),
      ).toBeVisible();
      await check(native.stages[0].cells);
      await page.screenshot({ path: `${root}/browser-source.png` });
    }
    await button(['edit', 'error'].includes(stage.stage) ? 'Data' : 'Calls').click();
    const target = ['edit', 'error'].includes(stage.stage)
      ? 'A2'
      : stage.stage === 'entry'
        ? 'B1'
        : 'C1';
    if (stage.stage === 'entry') await cell(target).click();
    else await cell(target).dblclick();
    await page
      .getByRole('textbox', {
        name: stage.stage === 'entry' ? 'Formula bar' : `Edit ${target}`,
        exact: true,
      })
      .fill(
        stage.stage === 'edit'
          ? '0'
          : stage.stage === 'error'
            ? '=1/0'
            : stage.stage === 'formula'
              ? '=SUMPRODUCT(IF(Data!A1:A3>6,Data!B1:B3,0))'
              : native.stages[0].cells.find((r: any) => r.ref === target).formula,
      );
    await page.keyboard.press(
      ['entry', 'formula'].includes(stage.stage) ? 'Control+Shift+Enter' : 'Enter',
    );
    if (['entry', 'formula'].includes(stage.stage))
      await expect(page.locator('.cell-address')).toHaveText(target);
    await check(stage.cells);
    await button('Undo').click();
    await check(native.stages[0].cells);
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
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${stage.stage}.xlsx`, bytes);
    exports[stage.stage] = hash(bytes);
    await upload(bytes, `Array reimport ${stage.stage}`);
    await check(stage.cells);
    await cell(stage.stage === 'ordinary' ? 'C1' : 'B1').click();
    const isArray = stage.cells.find(
      (r: any) => r.ref === (stage.stage === 'ordinary' ? 'C1' : 'B1'),
    ).hasArray;
    const bar = page.getByRole('textbox', { name: 'Formula bar', exact: true });
    if (isArray) await expect(bar).toHaveValue(/^\{=/);
    else await expect(bar).not.toHaveValue(/^\{=/);
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

test('legacy saved array metadata recovers while preserving edited cells and storage revision', async ({
  page,
}) => {
  const source = await fs.readFile('tests/fixtures/excel-array-contexts.xlsx');
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Legacy arrays.xlsx',
    buffer: source,
    mimeType: 'application/octet-stream',
  });
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const before = await page.evaluate(async () => {
    const r = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => (r.onsuccess = () => resolve(r.result)));
    const read = db.transaction('files').objectStore('files').getAll();
    const files = await new Promise<any[]>(
      (resolve) => (read.onsuccess = () => resolve(read.result)),
    );
    const file = files.find((f) => f.name === 'Legacy arrays');
    for (const s of file.content.sheets) delete s.arrayFormulas;
    file.content.sheets.find((s: any) => s.name === 'Data').cells.A2.value = '0';
    const tx = db.transaction('files', 'readwrite');
    tx.objectStore('files').put(file);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
    return { id: file.id, revision: file.revision };
  });
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Legacy arrays', exact: true }).click();
  await page.getByRole('button', { name: 'Calls', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'C1', exact: true })).toHaveText('8');
  const revision = await page.evaluate(async (id) => {
    const r = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => (r.onsuccess = () => resolve(r.result)));
    const read = db.transaction('files').objectStore('files').get(id);
    const file = await new Promise<any>((resolve) => (read.onsuccess = () => resolve(read.result)));
    db.close();
    return file.revision;
  }, before.id);
  expect(revision).toBe(before.revision);
});
