import { test, expect } from '@playwright/test';
import {
  Document,
  Packer,
  Paragraph,
  Header,
  ImageRun,
  FootnoteReferenceRun,
  PageOrientation,
} from 'docx';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';

test('Word retains source sections through editing, undo, reload, backup and native export', async ({
  page,
}) => {
  const root = '.local/word-sections';
  const backupLabel = 'Noffice backup Full editable model and retained original · .noffice';
  const download = async (label: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    return fs.readFile((await (await pending).path())!);
  };
  await fs.mkdir(root, { recursive: true });
  const header = (text: string) => new Header({ children: [new Paragraph(text)] });
  const input = await Packer.toBuffer(
    new Document({
      creator: 'Noffice test',
      evenAndOddHeaderAndFooters: true,
      styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
      footnotes: { 1: { children: [new Paragraph('Retained footnote.')] } },
      sections: [
        {
          properties: { page: { size: { width: 12240, height: 15840 } } },
          headers: { default: header('First odd'), even: header('First even') },
          children: [
            new Paragraph('Section one text.'),
            new Paragraph({
              children: [
                new ImageRun({
                  type: 'png',
                  data: Buffer.from(
                    'iVBORw0KGgoAAAANSUhEUgAAADwAAAAoCAIAAAAt2Q6oAAAAYklEQVR4nO3QUQ3AIAADUSB44QMjc4ACTE3B/GyCUNFLltwT0Fxa57pLxvfs0HIL7UYZTTGaYjTFaIrRFKMpRlOMphhN6bnp9xqh5V8+bTTFaIrRFKMpRlOMphhNMZpiNOUA1EcEkQvG4dcAAAAASUVORK5CYII=',
                    'base64',
                  ),
                  transformation: { width: 60, height: 40 },
                  floating: {
                    horizontalPosition: { offset: 508000 },
                    verticalPosition: { offset: 254000 },
                    wrap: { type: 1 },
                  },
                }),
              ],
            }),
          ],
        },
        {
          properties: {
            page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE } },
          },
          headers: { default: header('Second odd'), even: header('Second even') },
          children: [
            new Paragraph('Section two text.'),
            new Paragraph({ children: [new FootnoteReferenceRun(1)] }),
          ],
        },
      ],
    }),
  );
  await fs.writeFile(`${root}/source.docx`, input);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Sections.docx',
    buffer: input,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Sections',
  );
  const legacy = JSON.parse((await download(backupLabel)).toString());
  legacy.name = 'Migrated sections';
  delete legacy.content.docxStructure;
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Legacy-unchanged.noffice',
    buffer: Buffer.from(JSON.stringify(legacy)),
    mimeType: 'application/json',
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Migrated sections',
  );
  expect(await download('DOCX file Editable in Microsoft Word')).toEqual(input);
  await editor.locator('p').first().click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
  await page.keyboard.type('Section one edit.');
  await expect(editor.locator('p').first()).toHaveText('Section one edit.');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor.locator('p').first()).toHaveText('Section one text.');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(editor.locator('p').first()).toHaveText('Section one edit.');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Migrated sections', exact: true }).click();
  const backup = await download(backupLabel);
  const model = JSON.parse(backup.toString());
  expect(model.content.docxStructure.sections).toHaveLength(2);
  expect(model.content.docxStructure.sections[0].drawings).toHaveLength(1);
  expect(model.content.docxStructure.sections[1].notes).toHaveLength(1);
  // Exercise migration through a real legacy backup without losing the edited HTML.
  delete model.content.docxStructure;
  model.name = 'Migrated edited sections';
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Legacy.noffice',
    buffer: Buffer.from(JSON.stringify(model)),
    mimeType: 'application/json',
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Migrated edited sections',
  );
  await expect(editor.locator('p').first()).toHaveText('Section one edit.');
  const migrated = JSON.parse((await download(backupLabel)).toString());
  expect(migrated.content.docxStructure.sections).toHaveLength(2);
  const output = await download('DOCX file Editable in Microsoft Word');
  await fs.writeFile(`${root}/browser.docx`, output);
  const original = await JSZip.loadAsync(input),
    edited = await JSZip.loadAsync(output);
  for (const name of Object.keys(original.files).filter(
    (p) => p !== 'word/document.xml' && !original.files[p].dir,
  ))
    expect(await edited.file(name)!.async('uint8array')).toEqual(
      await original.file(name)!.async('uint8array'),
    );
  const xml = await edited.file('word/document.xml')!.async('string');
  expect(xml).toContain('Section one edit.');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Reimport.docx',
    buffer: output,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Reimport',
  );
  await expect(editor.locator('p').first()).toHaveText('Section one edit.');
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        browser: page.context().browser()!.version(),
        sourceHash: createHash('sha256').update(input).digest('hex'),
        exportHash: createHash('sha256').update(output).digest('hex'),
        workflowPassed: true,
      },
      null,
      2,
    ),
  );
});

test('legacy saved Word snapshots hydrate without replacing edits or writing a storage revision', async ({
  page,
}) => {
  const input = await Packer.toBuffer(
    new Document({ sections: [{ children: [new Paragraph('Original paragraph.')] }] }),
  );
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Legacy saved.docx',
    buffer: input,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Legacy saved',
  );
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const stored = (strip: boolean) =>
    page.evaluate(
      (strip) =>
        new Promise<{ revision: number; hasStructure: boolean }>((resolve, reject) => {
          const request = indexedDB.open('noffice-workspace');
          request.onerror = () => reject(request.error);
          request.onsuccess = () => {
            const db = request.result,
              tx = db.transaction('files', strip ? 'readwrite' : 'readonly'),
              store = tx.objectStore('files'),
              all = store.getAll();
            let result: { revision: number; hasStructure: boolean };
            all.onsuccess = () => {
              const file = all.result.find(
                (f: OfficeFile) => f.name === 'Legacy saved',
              ) as OfficeFile;
              if (file.content.kind !== 'word') {
                tx.abort();
                return;
              }
              if (strip) {
                file.content.html = file.content.html.replace('Original paragraph.', 'Saved edit.');
                delete file.content.docxStructure;
                store.put(file);
              }
              result = { revision: file.revision, hasStructure: !!file.content.docxStructure };
            };
            tx.oncomplete = () => {
              db.close();
              resolve(result);
            };
            tx.onabort = tx.onerror = () => {
              db.close();
              reject(tx.error);
            };
          };
        }),
      strip,
    );
  const before = await stored(true);
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Legacy saved', exact: true }).click();
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(editor.locator('p').first()).toHaveText('Saved edit.');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', {
      name: 'Noffice backup Full editable model and retained original · .noffice',
      exact: true,
    })
    .click();
  const backup = JSON.parse((await fs.readFile((await (await pending).path())!)).toString());
  expect(backup.content.docxStructure.sections).toHaveLength(1);
  expect(Buffer.from(backup.original.base64, 'base64')).toEqual(input);
  expect(await stored(false)).toEqual(before);
  await editor.locator('p').first().click();
  await page.keyboard.press('End');
  await page.keyboard.type(' Again.');
  await expect(editor.locator('p').first()).toHaveText('Saved edit. Again.');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const after = await stored(false);
  expect(after.revision).toBeGreaterThan(before.revision);
  expect(after.hasStructure).toBe(true);
});
