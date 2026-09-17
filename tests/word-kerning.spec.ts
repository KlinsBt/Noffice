import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import type { OfficeFile } from '../src/model';

const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-kerning.json', 'utf8'),
) as typeof import('./fixtures/native-word-kerning.json');

test('Word kerning thresholds survive size edits, history, reload and actual exports', async ({
  page,
}) => {
  test.setTimeout(90000);
  const root = '.local/word-kerning-acceptance';
  await fs.mkdir(root, { recursive: true });
  const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
  const source = await fs.readFile('tests/fixtures/word-kerning.docx');
  expect(hash(source)).toBe(reference.sourceSha256);
  const upload = async (data: Buffer, name: string) => {
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `${name}.docx`,
      buffer: data,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const values = () =>
    editor.locator('p').evaluateAll((paragraphs) =>
      paragraphs.map((p) => {
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        const node = walker.nextNode()!;
        const style = getComputedStyle(node.parentElement!);
        return {
          size: Math.round(parseFloat(style.fontSize) * 750) / 1000,
          kerning: style.fontKerning,
          text: p.textContent,
        };
      }),
    );
  const expected = reference.configuration.map((row) => ({
    size: row.size,
    kerning: row.threshold && row.size * 2 >= row.threshold ? 'normal' : 'none',
    text: '11 AV To',
  }));
  const assertValues = async (size: number) => {
    const rows = expected.map((row, i) =>
      i === 3 ? { ...row, size, kerning: size >= 12 ? 'normal' : 'none' } : row,
    );
    await expect.poll(values).toEqual(rows);
  };
  const download = async (name: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const data = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${name}.docx`, data);
    return data;
  };
  const print = async (name: string) =>
    hash(await page.pdf({ path: `${root}/${name}.pdf`, preferCSSPageSize: true }));
  await page.goto('/');
  await upload(source, 'Kerning source');
  await assertValues(10);
  const stages: {
    name: string;
    size: number;
    exportHash: string;
    printHash: string;
    values: Awaited<ReturnType<typeof values>>;
  }[] = [];
  const capture = async (name: string, size: number) => {
    await assertValues(size);
    const data = await download(name);
    stages.push({
      name,
      size,
      exportHash: hash(data),
      printHash: await print(name),
      values: await values(),
    });
    return data;
  };
  await capture('source', 10);
  let previous = 10;
  for (const size of [20, 10]) {
    await editor.focus();
    await page.keyboard.press('Control+Home');
    for (let i = 0; i < 3; i++) await page.keyboard.press('Control+ArrowDown');
    await page.keyboard.press('Home');
    await page.keyboard.press('Shift+End');
    await expect.poll(() => page.evaluate(() => getSelection()?.toString())).toBe('11 AV To');
    const input = page.getByRole('spinbutton', { name: 'Font size', exact: true });
    await input.fill(String(size));
    await input.press('Tab');
    await assertValues(size);
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await assertValues(previous);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await assertValues(size);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).first().click();
    const data = await capture(`size-${size}`, size);
    await upload(data, `Kerning ${size}`);
    await assertValues(size);
    previous = size;
  }
  // HTML paste is a supported authoring input. Exercise explicit zero in the
  // actual new-document writer, separately from retained-source patching.
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Word', exact: true }).click();
  await editor.focus();
  await editor.evaluate((el) => {
    const transfer = new DataTransfer();
    transfer.setData(
      'text/html',
      '<p><span data-word-kerning="0" style="font-family:Arial;font-size:10pt">11 AV To</span></p>',
    );
    el.dispatchEvent(
      new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: transfer }),
    );
  });
  await expect(editor.getByText('11 AV To', { exact: true })).toHaveCSS('font-kerning', 'none');
  const zero = await download('new-zero');
  await upload(zero, 'Zero reopened');
  await expect(editor.getByText('11 AV To', { exact: true })).toHaveCSS('font-kerning', 'none');
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceHash: hash(source),
        stages,
        newZeroHash: hash(zero),
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ) + '\n',
  );
});

for (const mixed of [false, true])
  test(`legacy kerning ${mixed ? 'rejects ambiguous edits and preserves a recoverable backup' : 'recovers edited text without a migration save'}`, async ({
    page,
  }) => {
    const name = mixed ? 'Ambiguous kerning' : 'Legacy kerning';
    const source = mixed
      ? await Packer.toBuffer(
          new Document({
            sections: [
              {
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({ text: 'AB', font: 'Arial', size: 20, kern: 16 }),
                      new TextRun({ text: 'CD', font: 'Arial', size: 20, kern: '0pt' }),
                    ],
                  }),
                ],
              },
            ],
          }),
        )
      : await fs.readFile('tests/fixtures/word-kerning.docx');
    await page.goto('/');
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({
        name: `${name}.docx`,
        buffer: source,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const stored = (strip: boolean) =>
      page.evaluate(
        ({ strip, name, mixed }) =>
          new Promise<{ revision: number; html: string; version?: number; original: number[] }>(
            (resolve, reject) => {
              const request = indexedDB.open('noffice-workspace');
              request.onerror = () => reject(request.error);
              request.onsuccess = () => {
                const db = request.result,
                  tx = db.transaction('files', strip ? 'readwrite' : 'readonly');
                const store = tx.objectStore('files'),
                  all = store.getAll();
                let result: {
                  revision: number;
                  html: string;
                  version?: number;
                  original: number[];
                };
                all.onsuccess = () => {
                  const file = all.result.find((f: OfficeFile) => f.name === name) as OfficeFile;
                  if (file.content.kind !== 'word') {
                    tx.abort();
                    return;
                  }
                  if (strip) {
                    delete file.content.kerningVersion;
                    const dom = new DOMParser().parseFromString(file.content.html, 'text/html');
                    for (const span of dom.querySelectorAll<HTMLElement>('[data-word-kerning]')) {
                      span.removeAttribute('data-word-kerning');
                      span.style.removeProperty('font-kerning');
                    }
                    dom.querySelectorAll('p')[mixed ? 0 : 2].append(' saved edit');
                    file.content.html = dom.body.innerHTML;
                    store.put(file);
                  }
                  result = {
                    revision: file.revision,
                    html: file.content.html,
                    version: file.content.kerningVersion,
                    original: [...new Uint8Array(file.original!.data)],
                  };
                };
                tx.oncomplete = () => {
                  db.close();
                  resolve(result);
                };
                tx.onerror = tx.onabort = () => {
                  db.close();
                  reject(tx.error);
                };
              };
            },
          ),
        { strip, name, mixed },
      );
    const before = await stored(true);
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).first().click();
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    if (mixed) {
      await expect(page.getByText(/edited text with missing mixed kerning metadata/)).toBeVisible();
      await expect(editor).toHaveCount(0);
    } else {
      await expect(editor.locator('p').nth(2)).toHaveText('11 AV To saved edit');
      await expect(
        editor.locator('p').nth(2).locator('[data-word-kerning="16"]').first(),
      ).toHaveCSS('font-kerning', 'normal');
    }
    expect(await stored(false)).toEqual(before);
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', {
        name: 'Noffice backup Full editable model and retained original · .noffice',
        exact: true,
      })
      .click();
    const backup = JSON.parse((await fs.readFile((await (await pending).path())!)).toString());
    expect(Buffer.from(backup.original.base64, 'base64')).toEqual(source);
    expect(backup.content.html).toContain(' saved edit');
    expect(backup.content.kerningVersion).toBe(mixed ? undefined : 1);
    if (mixed) expect(backup.content.html).toBe(before.html);
    else {
      await editor.locator('p').nth(2).click();
      await page.keyboard.press('End');
      await page.keyboard.insertText('!');
      await expect(editor.locator('p').nth(2)).toHaveText('11 AV To saved edit!');
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await expect(editor.locator('p').nth(2)).toHaveText('11 AV To saved edit');
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await expect(editor.locator('p').nth(2)).toHaveText('11 AV To saved edit!');
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      expect((await stored(false)).version).toBe(1);
      expect((await stored(false)).revision).toBeGreaterThan(before.revision);
    }
    const root = '.local/word-kerning-acceptance';
    await fs.mkdir(root, { recursive: true });
    await fs.writeFile(
      `${root}/legacy-${mixed ? 'failure' : 'recovery'}.json`,
      JSON.stringify(
        {
          passed: true,
          originalHash: createHash('sha256').update(source).digest('hex'),
          unchangedStorageThroughHydration: true,
          backupRetainsOriginalAndEdits: true,
          browser: page.context().browser()!.version(),
        },
        null,
        2,
      ) + '\n',
    );
  });
