import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
const root = '.local/excel-relative-names';
const native = JSON.parse(
  await fs.readFile('tests/fixtures/native-excel-relative-names.json', 'utf8'),
);
const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
test('relative names follow each caller through history, reload, errors and three actual XLSX exports', async ({
  page,
}) => {
  test.setTimeout(120000);
  const source = await fs.readFile('tests/fixtures/excel-relative-names.xlsx');
  expect(hash(source)).toBe(native.sourceSha256);
  const button = (name: string) => page.getByRole('button', { name, exact: true });
  const cell = (name: string) => page.getByRole('gridcell', { name, exact: true });
  const upload = async (buffer: Buffer, name: string) => {
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({ name: `${name}.xlsx`, buffer, mimeType: 'application/octet-stream' });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  const edit = async (sheet: string, ref: string, value: string) => {
    await button(sheet).click();
    await cell(ref).dblclick();
    await page.getByRole('textbox', { name: `Edit ${ref}`, exact: true }).fill(value);
    await page.keyboard.press('Enter');
  };
  const check = async (stage: string) => {
    const samples = native.stages.find((s: any) => s.stage === stage).cells;
    for (const name of ['Calls', 'Cost data']) {
      await button(name).click();
      for (const c of samples.filter((c: any) => c.sheet === name))
        await expect(cell(c.address)).toHaveText(
          c.valueType === 'System.Int32' ? '#DIV/0!' : c.value === null ? '' : String(c.value),
        );
    }
  };
  await page.goto('/');
  await upload(source, 'Relative names');
  await check('source');
  const exports: Record<string, string> = {};
  for (const stage of ['edit', 'delete', 'caller']) {
    if (stage !== 'caller') {
      await edit('Data', 'A3', stage === 'edit' ? '500' : '');
      await check(stage);
      await button('Undo').click();
      await check(stage === 'edit' ? 'source' : 'edit');
      await button('Redo').click();
      await check(stage);
    } else {
      await edit('Calls', 'D2', '=1/RowNext');
      await expect(cell('D2')).toHaveText('#DIV/0!');
      await button('Undo').click();
      await expect(cell('D2')).toHaveText('0');
      await button('Redo').click();
      await expect(cell('D2')).toHaveText('#DIV/0!');
      await edit('Calls', 'D3', '=RowNext');
      await check(stage);
    }
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const currentName = await page
      .getByRole('textbox', { name: 'File name', exact: true })
      .inputValue();
    await page.reload();
    await button('Recent files').click();
    await button(currentName).click();
    await check(stage);
    await button('Export').click();
    const pending = page.waitForEvent('download');
    await button('XLSX file Editable in Microsoft Excel').click();
    const output = await fs.readFile((await (await pending).path())!);
    exports[stage] = hash(output);
    await fs.mkdir(root, { recursive: true });
    await fs.writeFile(`${root}/${stage}.xlsx`, output);
    // Inspect native-typed formula caches before any native opening/recalculation.
    const zip = await JSZip.loadAsync(output);
    for (const name of ['Calls', 'Cost data']) {
      const xml = await zip
        .file(`xl/worksheets/sheet${name === 'Calls' ? 2 : 3}.xml`)!
        .async('string');
      const values = await page.evaluate((xml) => {
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        return Object.fromEntries(
          Array.from(doc.getElementsByTagNameNS('*', 'c')).map((c) => [
            c.getAttribute('r'),
            {
              value: c.getElementsByTagNameNS('*', 'v')[0]?.textContent,
              type: c.getAttribute('t') || 'n',
              formula: c.getElementsByTagNameNS('*', 'f')[0]?.textContent,
            },
          ]),
        );
      }, xml);
      for (const c of native.stages
        .find((s: any) => s.stage === stage)
        .cells.filter((c: any) => c.sheet === name && c.formula.startsWith('='))) {
        expect(values[c.address]).toEqual({
          value: c.valueType === 'System.Int32' ? '#DIV/0!' : String(c.value),
          type: c.valueType === 'System.Int32' ? 'e' : 'n',
          formula: c.formula.slice(1),
        });
      }
    }
    await upload(output, `Relative ${stage}`);
    await check(stage);
  }
  // Recover from the actual exported error, then undo back to that same error.
  await edit('Data', 'A3', '7');
  await button('Calls').click();
  await expect(cell('D2')).not.toHaveText('#DIV/0!');
  await button('Undo').click();
  await expect(cell('D2')).toHaveText('#DIV/0!');
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
test('legacy relative-name origins hydrate without changing saved revisions or edited cells', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles('tests/fixtures/excel-relative-names.xlsx');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const original = await page.evaluate(async () => {
    const request = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => {
      request.onsuccess = () => resolve(request.result);
    });
    const query = db.transaction('files').objectStore('files').getAll();
    const files = await new Promise<any[]>((resolve) => {
      query.onsuccess = () => resolve(query.result);
    });
    const file = files.find((f) => f.name === 'excel-relative-names');
    for (const sheet of file.content.sheets)
      for (const definition of sheet.nameDefinitions) delete definition.referenceOrigin;
    file.content.sheets[0].cells.A3.value = '500';
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
  await page.getByRole('button', { name: 'excel-relative-names', exact: true }).click();
  await page.getByRole('button', { name: 'Calls', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'D2', exact: true })).toHaveText('500');
  const revision = await page.evaluate(async (id) => {
    const request = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => {
      request.onsuccess = () => resolve(request.result);
    });
    const query = db.transaction('files').objectStore('files').get(id);
    const file = await new Promise<any>((resolve) => {
      query.onsuccess = () => resolve(query.result);
    });
    db.close();
    return file.revision;
  }, original.id);
  expect(revision).toBe(original.revision);
});
