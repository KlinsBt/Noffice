import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

for (const failure of ['late-abort', 'other-tab'])
  test(`Scoped section geometry survives ${failure}`, async ({ page, context }) => {
    const run = process.env.NOFFICE_SCOPED_SECTION_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Invalid recovery run');
    const root = `.local/word-section-scoped-layout/recovery-${run}/${failure}`;
    await fs.mkdir(root, { recursive: true });
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: 'Section recovery.docx',
      buffer: await fs.readFile('tests/fixtures/word-boundary-delete/paragraph-period.docx'),
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
    await editor.evaluate((el) =>
      (
        el as HTMLElement & { editor: import('@tiptap/core').Editor }
      ).editor.commands.setTextSelection(4),
    );
    await page.getByRole('button', { name: 'Layout', exact: true }).click();
    await page
      .getByRole('combobox', { name: 'Apply page setup to', exact: true })
      .selectOption('section');
    await page
      .getByRole('combobox', { name: 'Page orientation', exact: true })
      .selectOption('landscape');
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
    ).toHaveLength(1);
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
    await page.getByRole('button', { name: 'Layout', exact: true }).click();
    if (other) {
      await expect(
        page.getByRole('combobox', { name: 'Page orientation', exact: true }),
      ).toHaveValue('portrait');
      await page.goto('/');
      await page.locator('input[type=file][multiple]').setInputFiles(root + '/unsaved.docx');
      await waitWordLayout(editor);
      await other.close();
      await page.getByRole('button', { name: 'Layout', exact: true }).click();
    }
    await expect(page.getByRole('combobox', { name: 'Page orientation', exact: true })).toHaveValue(
      'landscape',
    );
    await fs.writeFile(
      root + '/report.json',
      JSON.stringify(
        {
          failure,
          buildHash: await wordStoryBuildHash(),
          exportHash: createHash('sha256').update(bytes).digest('hex'),
          testHash: createHash('sha256')
            .update(await fs.readFile('tests/word-section-scoped-recovery.spec.ts'))
            .digest('hex'),
        },
        null,
        2,
      ),
    );
  });
