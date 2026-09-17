import { test, expect } from '@playwright/test';

test('an import made while the initial library loads survives initialization, editing and reload', async ({
  page,
}) => {
  await page.addInitScript(() => {
    const original = IDBObjectStore.prototype.getAll;
    const descriptor = Object.getOwnPropertyDescriptor(IDBRequest.prototype, 'onsuccess')!;
    let first = true;
    IDBObjectStore.prototype.getAll = function (...args) {
      const request = original.apply(this, args);
      if (this.name === 'files' && first) {
        first = false;
        Object.defineProperty(request, 'onsuccess', {
          set(handler) {
            descriptor.set!.call(request, function (event: Event) {
              Object.assign(window, { releaseLibrary: () => handler.call(request, event) });
            });
          },
        });
      }
      return request;
    };
  });
  const ready = () =>
    page.evaluate(
      () =>
        typeof (window as unknown as { releaseLibrary?: () => void }).releaseLibrary === 'function',
    );
  const release = () =>
    page.evaluate(() => (window as unknown as { releaseLibrary: () => void }).releaseLibrary());
  await page.goto('/');
  await expect.poll(ready).toBe(true);
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Early import.txt',
    mimeType: 'text/plain',
    buffer: Buffer.from('Imported before the library finished'),
  });
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(editor).toContainText('Imported before the library finished');
  await editor.click();
  await page.keyboard.press('Control+End');
  await page.keyboard.type(' and edited');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await release();
  await expect(page.getByRole('application', { name: 'Noffice workspace' })).toHaveAttribute(
    'aria-busy',
    'false',
  );
  await expect(editor).toContainText('Imported before the library finished and edited');
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Early import',
  );
  await page.reload();
  await expect.poll(ready).toBe(true);
  await release();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Early import', exact: true }).click();
  await expect(editor).toContainText('Imported before the library finished and edited');
  await expect(page.getByText('Not saved', { exact: true })).toHaveCount(0);
});
