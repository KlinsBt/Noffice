import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { waitWordLayout, wordPlacements } from './word-pagination-helpers';

test('Word legacy unformatted paragraphs inherit defaults without rewriting saved content', async ({
  page,
}) => {
  const root = '.local/word-default-inheritance';
  await fs.mkdir(root, { recursive: true });
  const labels = Array.from({ length: 70 }, (_, i) => `line${String(i + 1).padStart(2, '0')}`);
  const html = `<p>${labels.join('<br>')}</p>`;
  const reference = JSON.parse(
    readFileSync('tests/fixtures/native-word-inter-long-metrics.json', 'utf8'),
  );
  // The independent 80-line paragraph has the same font, leading and body
  // rectangle. Its first 70 line assignments also cover this shorter paragraph.
  const expectedPages = reference.cases
    .find((c: { name: string }) => c.name === '12-default')
    .starts.slice(0, 70)
    .map((line: { page: number }) => line.page);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Inherited defaults.noffice',
    mimeType: 'application/json',
    buffer: Buffer.from(
      JSON.stringify({
        format: 'noffice',
        version: 1,
        name: 'Inherited defaults',
        content: { kind: 'word', html, paper: 'a4', margin: 'normal' },
      }),
    ),
  });
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Inherited defaults',
  );
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const stored = () =>
    page.evaluate(
      () =>
        new Promise<{ html: string; revision: number }>((resolve, reject) => {
          const request = indexedDB.open('noffice-workspace');
          request.onerror = () => reject(request.error);
          request.onsuccess = () => {
            const db = request.result,
              tx = db.transaction('files'),
              all = tx.objectStore('files').getAll();
            let value: { html: string; revision: number };
            all.onsuccess = () => {
              const file = all.result.find((f) => f.name === 'Inherited defaults');
              value = { html: file.content.html, revision: file.revision };
            };
            tx.oncomplete = () => {
              db.close();
              resolve(value);
            };
            tx.onerror = tx.onabort = () => {
              db.close();
              reject(tx.error);
            };
          };
        }),
    );
  const before = await stored();
  await expect(page.locator('.section-page')).toHaveCount(3);
  await waitWordLayout(editor);
  const source = await wordPlacements(editor);
  expect(
    source[0].filter((c, i) => i === 0 || source[0][i - 1].text === '\n').map((c) => c.page),
  ).toEqual(expectedPages);
  expect(await stored()).toEqual(before);
  const outputs: Record<string, string> = {};
  const download = async (stage: string, kind: 'docx' | 'pdf') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', {
        name: kind === 'pdf' ? 'PDF file' : 'DOCX file Editable in Microsoft Word',
        exact: true,
      })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${stage}.${kind}`, bytes);
    outputs[`${stage}.${kind}`] = createHash('sha256').update(bytes).digest('hex');
  };
  await download('source', 'docx');
  await download('source', 'pdf');
  await editor.focus();
  await page.keyboard.press('Control+End');
  await page.keyboard.insertText(' added');
  await expect(editor).toContainText('line70 added');
  await waitWordLayout(editor);
  const edited = await wordPlacements(editor);
  await download('edited', 'docx');
  await download('edited', 'pdf');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor).not.toContainText(' added');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(editor).toContainText('line70 added');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Inherited defaults', exact: true }).click();
  await expect.poll(() => wordPlacements(editor)).toEqual(source);
  await download('restored', 'docx');
  await download('restored', 'pdf');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Inherited reimport.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await fs.readFile(`${root}/restored.docx`),
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Inherited reimport',
  );
  await expect.poll(() => wordPlacements(editor)).toEqual(source);
  const reimported = await wordPlacements(editor);
  await download('reimported', 'docx');
  await download('reimported', 'pdf');
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        before,
        source,
        edited,
        restored: source,
        reimported,
        outputs,
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ),
  );
});
