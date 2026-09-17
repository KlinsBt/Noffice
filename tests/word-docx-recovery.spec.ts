import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';

test('Word rejects edits during DOCX generation and retries without an extra undo event', async ({
  page,
}) => {
  await page.addInitScript(() => {
    const state = window as Window & {
      holdDocx?: boolean;
      docxWaiting?: boolean;
      releaseDocx?: () => void;
    };
    const digest = crypto.subtle.digest.bind(crypto.subtle);
    crypto.subtle.digest = async (algorithm, data) => {
      // Delay the platform operation that begins new-package save-session
      // generation. No editor or application command is replaced by the test.
      if (
        state.holdDocx &&
        algorithm === 'SHA-256' &&
        /^[a-f\d]{8}(?:-[a-f\d]{4}){3}-[a-f\d]{12}$/i.test(new TextDecoder().decode(data))
      ) {
        state.holdDocx = false;
        state.docxWaiting = true;
        await new Promise<void>((resolve) => {
          state.releaseDocx = resolve;
        });
      }
      return digest(algorithm, data);
    };
  });
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Word', exact: true }).click();
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await editor.fill('Original');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await page.evaluate(() => {
    (window as Window & { holdDocx?: boolean }).holdDocx = true;
  });
  const start = async () => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
  };
  await start();
  await expect
    .poll(() => page.evaluate(() => !!(window as Window & { docxWaiting?: boolean }).docxWaiting))
    .toBe(true);
  await page.keyboard.press('Escape');
  await editor.focus();
  await page.keyboard.press('Control+End');
  await page.keyboard.insertText(' newer');
  await expect(editor).toContainText('Original newer');
  await page.evaluate(() => {
    (window as Window & { releaseDocx?: () => void }).releaseDocx?.();
  });
  await expect(page.getByRole('alert')).toContainText('Document changed before export');
  expect(downloads).toEqual([]);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor).toHaveText('Original');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(editor).toHaveText('Original newer');
  const pending = page.waitForEvent('download');
  await start();
  const bytes = await fs.readFile((await (await pending).path())!);
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file('word/document.xml')!.async('string');
  expect([...xml.matchAll(/<w:t(?:\s[^>]*)?>(.*?)<\/w:t>/g)].map((m) => m[1]).join('')).toBe(
    'Original newer',
  );
  expect(downloads).toHaveLength(1);
  // Saving does not add a history event: the next Undo still removes typing.
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor).toHaveText('Original');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(editor).toHaveText('Original newer');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Untitled document', exact: true }).click();
  await expect(editor).toHaveText('Original newer');
  await fs.mkdir('.local/word-docx-recovery', { recursive: true });
  await fs.writeFile('.local/word-docx-recovery/retried.docx', bytes);
});
