import { test, expect } from '@playwright/test';
import { clickWordText, waitWordLayout } from './word-pagination-helpers';

test('Word layout refresh preserves a native End caret before selectionchange is delivered', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1500, height: 1200 });
  await page.goto('/');
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles('tests/fixtures/word-section-flow-oddPage-lines-31.docx');
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(editor.locator('p').first()).toContainText('s1l31');
  await clickWordText(page, editor, 's1l31');
  await waitWordLayout(editor);
  // Delay only delivery of the native selection notification, as can happen
  // when a pending layout frame precedes ProseMirror's selection observer.
  await page.evaluate(() => {
    const block = (e: Event) => e.stopImmediatePropagation();
    window.addEventListener('selectionchange', block, true);
    (window as unknown as { releaseSelection: () => void }).releaseSelection = () =>
      window.removeEventListener('selectionchange', block, true);
  });
  await page.keyboard.press('End');
  const caret = () =>
    editor.evaluate((root) => {
      const selection = document.getSelection()!;
      const range = document.createRange();
      range.selectNodeContents(root.querySelector('p')!);
      range.setEnd(selection.focusNode!, selection.focusOffset);
      return range.toString();
    });
  await expect.poll(caret).toMatch(/s1l31$/);
  await editor.evaluate((root) => {
    (root.closest('.paper-wrap') as HTMLElement).style.zoom = '1.25';
    root.dispatchEvent(new Event('wordlayoutchange'));
  });
  await waitWordLayout(editor);
  await expect.poll(caret).toMatch(/s1l31$/);
  await page.evaluate(() =>
    (window as unknown as { releaseSelection: () => void }).releaseSelection(),
  );
  await page.keyboard.press('Shift+Enter');
  await page.keyboard.insertText('added');
  await expect(editor.locator('p').first()).toContainText('s1l31added');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor.locator('p').first()).not.toContainText('added');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(editor.locator('p').first()).toContainText('s1l31added');
});
