import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { PDFDocument } from 'pdf-lib';
import reference from './fixtures/word-first-section-parity/reference.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
for (const sample of reference.rows)
  test(`First section parity ${sample.name}: editing, history, reload and files`, async ({
    page,
  }) => {
    const run = process.env.NOFFICE_FIRST_PARITY_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Use versioned first-section evidence.');
    const root = `.local/word-section-save-controls/${run}/${sample.name}`;
    await fs.mkdir(root, { recursive: true });
    await page.context().grantPermissions(['local-fonts']);
    await page.goto('/');
    const source = await fs.readFile(sample.source);
    expect(hash(source)).toBe(sample.sourceHash);
    await page.locator('input[type=file][multiple]').setInputFiles(sample.source);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    await waitWordLayout(editor);
    const text = () =>
      editor.evaluate((el) => {
        const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
        return e.state.doc.textBetween(0, e.state.doc.content.size, '\r') + '\r';
      });
    const original = sample.native.text.replaceAll('\f', '\r');
    expect(await text()).toBe(original);
    await editor.focus();
    await page.keyboard.press('Control+End');
    await page.keyboard.type('X');
    const edited = original.slice(0, -1) + 'X\r';
    await expect.poll(text).toBe(edited);
    await page.keyboard.press('Control+z');
    await expect.poll(text).toBe(original);
    await page.keyboard.press('Control+y');
    await expect.poll(text).toBe(edited);
    await page.keyboard.press('Control+z');
    await expect.poll(text).toBe(original);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const title = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: title, exact: true }).click();
    await waitWordLayout(editor);
    expect(await text()).toBe(original);
    const hashes: Record<string, string> = {};
    for (const [ext, label] of [
      ['docx', 'DOCX file Editable in Microsoft Word'],
      ['pdf', 'PDF file'],
    ]) {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const waiting = page.waitForEvent('download');
      await page.getByRole('button', { name: label, exact: true }).click();
      const path = root + '/reloaded.' + ext;
      await (await waiting).saveAs(path);
      const bytes = await fs.readFile(path);
      hashes[ext] = hash(bytes);
      if (ext === 'docx') expect(hash(bytes)).toBe(sample.sourceHash);
      else expect((await PDFDocument.load(bytes)).getPageCount()).toBe(sample.native.pages);
    }
    await fs.writeFile(
      root + '/report.json',
      JSON.stringify(
        {
          name: sample.name,
          sourceHash: sample.sourceHash,
          hashes,
          buildHash: await wordStoryBuildHash(),
          testHash: hash(await fs.readFile('tests/word-first-section-parity.spec.ts')),
          referenceHash: hash(
            await fs.readFile('tests/fixtures/word-first-section-parity/reference.json'),
          ),
          scope:
            'Native first-section parity controls through real editing/history/reload, original DOCX preservation and actual PDF page count. Header/text/glyph comparison is independent.',
        },
        null,
        2,
      ),
    );
  });
