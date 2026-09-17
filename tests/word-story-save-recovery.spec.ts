import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { wordStoryBuildHash } from './word-story-artifacts';

async function options(page: Page, first: boolean) {
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: 'Page options', exact: true }).click();
  await page.getByRole('checkbox', { name: 'Different first page', exact: true }).setChecked(first);
  await page.getByRole('button', { name: 'Apply page options', exact: true }).click();
}
async function stored(page: Page) {
  return page.evaluate(async () => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const r = indexedDB.open('noffice-workspace', 1);
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
    try {
      return await new Promise<{ revision: number; name: string; content: unknown }>(
        (resolve, reject) => {
          const r = db.transaction('files').objectStore('files').getAll();
          r.onsuccess = () => resolve(r.result.find((f) => f.name === 'Recovery'));
          r.onerror = () => reject(r.error);
        },
      );
    } finally {
      db.close();
    }
  });
}

for (const failure of ['late-abort', 'other-tab'] as const)
  test(`Word story options preserve unsaved work and recover from ${failure}`, async ({
    page,
    context,
  }) => {
    test.setTimeout(90000);
    await context.grantPermissions(['local-fonts']);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const root = `.local/word-story-options/recovery/${failure}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile('tests/fixtures/word-section-stories/first-even.docx');
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: 'Recovery.docx',
      buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.locator('.section-page')).toHaveCount(6);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const original = await stored(page);
    let other: Page | undefined;
    if (failure === 'late-abort') {
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
          restoreStorySave: () => {
            IDBObjectStore.prototype.put = original;
          },
        });
      });
    } else {
      other = await context.newPage();
      await other.goto('/');
      await other.getByRole('button', { name: 'Recent files', exact: true }).click();
      await other.getByRole('button', { name: 'Recovery', exact: true }).click();
      await other.getByRole('textbox', { name: 'File name', exact: true }).fill('Recovery newer');
      await expect(other.getByText('Saved on this device', { exact: true })).toBeVisible();
    }
    await options(page, false);
    await expect(page.getByText('Not saved', { exact: true })).toBeVisible();
    await expect(page.getByRole('alert')).toBeVisible();
    if (failure === 'other-tab') await expect(page.getByRole('alert')).toContainText('another tab');
    else expect(await stored(page)).toEqual(original);

    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const waiting = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    await (await waiting).saveAs(`${root}/unsaved.docx`);
    const bytes = await fs.readFile(`${root}/unsaved.docx`);
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')!.async('string');
    expect((xml.match(/<w:titlePg(?:\s|\/|>)/g) || []).length).toBe(1);
    expect(await fs.readFile('tests/fixtures/word-section-stories/first-even.docx')).toEqual(
      source,
    );
    if (failure === 'late-abort') {
      await page.evaluate(() =>
        (window as unknown as { restoreStorySave: () => void }).restoreStorySave(),
      );
      await page.keyboard.press('Control+s');
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      await expect(page.getByRole('alert')).toHaveCount(0);
      expect((await stored(page)).revision).toBe(original.revision + 1);
      await page.getByRole('button', { name: 'Home', exact: true }).click();
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    }
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page
      .getByRole('button', {
        name: failure === 'late-abort' ? 'Recovery' : 'Recovery newer',
        exact: true,
      })
      .click();
    if (other) {
      await expect(page.locator('.word-page-story').first()).toHaveText('S1 header first');
      await page.locator('input[type=file][multiple]').setInputFiles(`${root}/unsaved.docx`);
      await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
        'unsaved',
      );
      await other.close();
    }
    await expect(page.locator('.word-page-story').first()).toHaveText('S1 header default');
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify(
        {
          failure,
          errors,
          sourceHash: createHash('sha256').update(source).digest('hex'),
          exportHash: createHash('sha256').update(bytes).digest('hex'),
          buildHash: await wordStoryBuildHash(),
        },
        null,
        2,
      ),
    );
  });
