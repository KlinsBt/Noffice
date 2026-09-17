import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';

test('Word replaces a focused editor without a render-time state mutation or losing saved edits', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  const bytes = await fs.readFile('tests/fixtures/word-inline-flow-column-midline-atLeast.docx');
  await page.goto('/');
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  for (const name of ['First focused document', 'Second focused document']) {
    // File input replacement removes the currently focused editor as part of
    // Svelte's keyed render. Its native blur must not mutate ribbon state there.
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `${name}.docx`,
      buffer: bytes,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
    await editor.focus();
    await page.keyboard.press('Control+End');
    await page.keyboard.insertText(' saved');
    await expect(editor).toContainText('after saved');
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  }
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'First focused document', exact: true }).click();
  await expect(editor).toContainText('after saved');
  await editor.focus();
  await page.keyboard.press('Control+End');
  await page.keyboard.insertText(' again');
  await expect(editor).toContainText('after saved again');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor).toContainText('after saved');
  await expect(editor).not.toContainText('again');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(editor).toContainText('after saved again');
  expect(errors).toEqual([]);
});
