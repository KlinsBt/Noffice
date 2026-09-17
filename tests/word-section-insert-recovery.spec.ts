import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

for (const fixture of ['paragraph', 'numbered', 'bulleted'])
for (const failure of ['late-abort', 'other-tab'])
  test(`Inserted sections${fixture === 'paragraph' ? '' : ` in ${fixture} lists`} survive ${failure}`, async ({ page, context }) => {
    const run = process.env.NOFFICE_SECTION_INSERT_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Invalid recovery run');
    const root = `.local/word-section-insertion/recovery-${run}/${failure}${fixture === 'paragraph' ? '' : `-${fixture}`}`;
    await fs.mkdir(root, { recursive: true });
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: 'Section recovery.docx',
      buffer: await fs.readFile(fixture === 'paragraph'
        ? 'tests/fixtures/word-boundary-delete/paragraph-period.docx'
        : `tests/fixtures/word-section-containers/${fixture}.docx`),
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    await waitWordLayout(editor);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const stored = () =>
      page.evaluate(async () => {
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          const r = indexedDB.open('noffice-workspace', 1);
          r.onsuccess = () => resolve(r.result);
          r.onerror = () => reject(r.error);
        });
        try {
          return await new Promise<unknown[]>((resolve, reject) => {
            const r = db.transaction('files').objectStore('files').getAll();
            r.onsuccess = () => resolve(r.result);
            r.onerror = () => reject(r.error);
          });
        } finally {
          db.close();
        }
      });
    const original = await stored();
    let other: import('@playwright/test').Page | undefined;
    if (failure === 'late-abort')
      await page.evaluate(() => {
        const original = IDBObjectStore.prototype.put;
        IDBObjectStore.prototype.put = function (...args: Parameters<typeof original>) {
          const request = original.apply(this, args);
          if (this.name === 'files') {
            const tx = this.transaction;
            request.addEventListener('success', () => tx.abort(), { once: true });
          }
          return request;
        };
        Object.assign(window, {
          restoreSectionSave: () => {
            IDBObjectStore.prototype.put = original;
          },
        });
      });
    else {
      other = await context.newPage();
      await other.goto('/');
      await other.getByRole('button', { name: 'Recent files', exact: true }).click();
      await other.getByRole('button', { name: 'Section recovery', exact: true }).click();
      await other
        .getByRole('textbox', { name: 'File name', exact: true })
        .fill('Section recovery newer');
      await expect(other.getByText('Saved on this device', { exact: true })).toBeVisible();
    }
    const baseline = await stored();
    await editor.evaluate((el, fixture) => {
      const editor = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
      let position = 4;
      if (fixture !== 'paragraph') editor.state.doc.descendants((node, from) => {
        if (node.attrs.sourceParagraph === '0:2') position = from + 4;
      });
      editor.commands.setTextSelection(position);
    }, fixture);
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page
      .getByRole('combobox', { name: 'Insert section break', exact: true })
      .selectOption('nextPage');
    await expect(page.getByText('Not saved', { exact: true })).toBeVisible();
    await expect(page.getByRole('alert')).toBeVisible();
    if (failure === 'other-tab') await expect(page.getByRole('alert')).toContainText('another tab');
    expect(await stored()).toEqual(baseline);
    if (failure === 'late-abort') expect(baseline).toEqual(original);
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const waiting = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    await (await waiting).saveAs(root + '/unsaved.docx');
    const bytes = await fs.readFile(root + '/unsaved.docx'),
      zip = await JSZip.loadAsync(bytes);
    expect(
      (await zip.file('word/document.xml')!.async('string')).match(/<w:sectPr(?:\s|>)/g) || [],
    ).toHaveLength(2);
    if (failure === 'late-abort') {
      await page.evaluate(() =>
        (window as unknown as { restoreSectionSave: () => void }).restoreSectionSave(),
      );
      await page.keyboard.press('Control+s');
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    }
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page
      .getByRole('button', {
        name: failure === 'late-abort' ? 'Section recovery' : 'Section recovery newer',
        exact: true,
      })
      .click();
    await waitWordLayout(editor);
    if (other) {
      expect(
        await editor.evaluate(
          (el) =>
            (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.state.doc.attrs
              .wordSectionState.breaks.length,
        ),
      ).toBe(0);
      await page.locator('input[type=file][multiple]').setInputFiles(root + '/unsaved.docx');
      await waitWordLayout(editor);
      await other.close();
    }
    expect(
      await editor.evaluate(
        (el) =>
          (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.state.doc.attrs
            .wordSectionState.breaks.length,
      ),
    ).toBe(1);
    await fs.writeFile(
      root + '/report.json',
      JSON.stringify(
        {
          failure,
          buildHash: await wordStoryBuildHash(),
          exportHash: createHash('sha256').update(bytes).digest('hex'),
          testHash: createHash('sha256')
            .update(await fs.readFile('tests/word-section-insert-recovery.spec.ts'))
            .digest('hex'),
        },
        null,
        2,
      ),
    );
  });
