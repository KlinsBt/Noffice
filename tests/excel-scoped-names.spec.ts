import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
const root = '.local/excel-scoped-names';
const native = JSON.parse(
  await fs.readFile('tests/fixtures/native-excel-scoped-names.json', 'utf8'),
);
test('scoped names calculate through edit, undo, redo, reload and native XLSX export', async ({
  page,
}) => {
  const source = await fs.readFile('tests/fixtures/excel-scoped-names.xlsx');
  const tab = (name: string) => page.getByRole('button', { name, exact: true });
  const cell = (name: string) => page.getByRole('gridcell', { name, exact: true });
  const upload = async (buffer: Buffer, name: string) => {
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({ name: `${name}.xlsx`, buffer, mimeType: 'application/octet-stream' });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  await page.goto('/');
  await upload(source, 'Scoped names');
  for (const sheet of ['Data', 'Cost data', 'Summary']) {
    await tab(sheet).click();
    for (const item of native.cells.filter((c: any) => c.sheet === sheet))
      await expect(cell(item.address)).toHaveText(String(item.value));
  }
  await tab('Data').click();
  await cell('A1').dblclick();
  await page.getByRole('textbox', { name: 'Edit A1', exact: true }).fill('5');
  await page.keyboard.press('Enter');
  await expect(cell('C1')).toHaveText('15');
  await expect(cell('B1')).toHaveText('11');
  await tab('Undo').click();
  await expect(cell('C1')).toHaveText('6');
  await tab('Redo').click();
  await expect(cell('C1')).toHaveText('15');
  await tab('Cost data').click();
  await expect(cell('B1')).toHaveText('7');
  await expect(cell('E1')).toHaveText('5');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await tab('Recent files').click();
  await tab('Scoped names').click();
  await tab('Summary').click();
  await expect(cell('C1')).toHaveText('15');
  await expect(cell('F1')).toHaveText('20');
  await tab('Export').click();
  const pending = page.waitForEvent('download');
  await tab('XLSX file Editable in Microsoft Excel').click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.mkdir(root, { recursive: true });
  await fs.writeFile(`${root}/browser.xlsx`, output);
  await upload(output, 'Names result');
  await tab('Cost data').click();
  await expect(cell('B1')).toHaveText('7');
  await expect(cell('C1')).toHaveText('15');
  await expect(cell('E1')).toHaveText('5');
  await tab('Data').click();
  await cell('A1').dblclick();
  await page.getByRole('textbox', { name: 'Edit A1', exact: true }).fill('');
  await page.keyboard.press('Enter');
  await expect(cell('C1')).toHaveText('0');
  await tab('Undo').click();
  await expect(cell('C1')).toHaveText('15');
  await tab('Redo').click();
  await expect(cell('C1')).toHaveText('0');
  await tab('Export').click();
  const deletedDownload = page.waitForEvent('download');
  await tab('XLSX file Editable in Microsoft Excel').click();
  const deleted = await fs.readFile((await (await deletedDownload).path())!);
  await fs.writeFile(`${root}/deleted.xlsx`, deleted);
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceSha256: hash(source),
        exportSha256: hash(output),
        deleteSha256: hash(deleted),
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ),
  );
});
test('legacy saved scoped names recover their identities while retaining edits and storage revision', async ({
  page,
}) => {
  const source = await fs.readFile('tests/fixtures/excel-scoped-names.xlsx');
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Legacy names.xlsx',
    buffer: source,
    mimeType: 'application/octet-stream',
  });
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const before = await page.evaluate(async () => {
    const r = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => {
      r.onsuccess = () => resolve(r.result);
    });
    const read = db.transaction('files').objectStore('files').getAll();
    const files = await new Promise<any[]>((resolve) => {
      read.onsuccess = () => resolve(read.result);
    });
    const file = files.find((f) => f.name === 'Legacy names');
    for (const s of file.content.sheets) delete s.nameDefinitions;
    file.content.sheets[0].cells.A1.value = '5';
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
  await page.getByRole('button', { name: 'Legacy names', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'C1', exact: true })).toHaveText('15');
  const revision = await page.evaluate(async (id) => {
    const r = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => {
      r.onsuccess = () => resolve(r.result);
    });
    const read = db.transaction('files').objectStore('files').get(id);
    const file = await new Promise<any>((resolve) => {
      read.onsuccess = () => resolve(read.result);
    });
    db.close();
    return file.revision;
  }, before.id);
  expect(revision).toBe(before.revision);
});
