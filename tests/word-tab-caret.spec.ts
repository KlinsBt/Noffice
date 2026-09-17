import { test, expect } from '@playwright/test';

test('tab repaint preserves a native caret move awaiting selectionchange delivery', async ({ page }) => {
  await page.addInitScript(() => {
    const probe = window as unknown as { focusSettled?: Promise<void> };
    const schedule = window.setTimeout.bind(window);
    window.setTimeout = ((handler: TimerHandler, delay?: number, ...args: unknown[]) => {
      // Observe completion of ProseMirror's real delayed focus restoration.
      // Holding selectionchange before that unrelated callback runs would
      // deliberately prevent focus from settling, rather than isolate repaint.
      if (typeof handler === 'function' && String(handler).includes('domSelectionRange') &&
          String(handler).includes('currentSelection') && document.activeElement?.getAttribute('aria-label') === 'Header or footer text') {
        let finish!: () => void;probe.focusSettled = new Promise<void>((resolve) => { finish = resolve; });
        return schedule(() => { try { handler(...args); } finally { finish(); } }, delay);
      }
      return schedule(handler, delay, ...args);
    }) as typeof window.setTimeout;
    document.addEventListener('selectionchange', (event) => {
      const blocked = (window as unknown as { holdTabSelection?: boolean }).holdTabSelection;
      const node = document.getSelection()?.focusNode;
      if (blocked && node?.parentElement?.closest('[aria-label="Header or footer text"]')) event.stopImmediatePropagation();
    }, true);
  });
  await page.setViewportSize({ width: 1500, height: 1200 });await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles('tests/fixtures/word-tab-stops-stories.docx');
  await expect(page.getByRole('textbox', { name: 'Document text', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Insert', exact: true }).click();await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
  const story = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  const tabs = story.locator('[data-word-tab-measured=true]');await expect(tabs).toHaveCount(2);
  const initialWidth = await tabs.last().evaluate((tab) => tab.getBoundingClientRect().width);
  await story.focus();
  await page.waitForFunction(() => !!(window as unknown as { focusSettled?: Promise<void> }).focusSettled);
  await page.evaluate(() => (window as unknown as { focusSettled: Promise<void> }).focusSettled);
  await page.keyboard.press('Control+Home');
  await page.evaluate(() => { (window as unknown as { holdTabSelection: boolean }).holdTabSelection = true; });
  // Hold only the browser's delayed notification, leaving actual native arrow
  // movement intact. No keyup is sent until after the asynchronous repaint.
  await page.keyboard.down('ArrowRight');
  const selection = () => story.evaluate((host) => {
    const editor = (host as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    const dom = document.getSelection()!;
    return { model: editor.state.selection.$head.parentOffset,
      native: editor.state.doc.resolve(editor.view.posAtDOM(dom.focusNode!, dom.focusOffset)).parentOffset };
  });
  expect(await selection()).toEqual({ model: 0, native: 1 });
  await story.evaluate((host) => {
    const editor = (host as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    (window as unknown as { tabRepaint: Promise<void> }).tabRepaint = new Promise<void>((resolve) => {
      const painted = ({ transaction }: { transaction: import('@tiptap/pm/state').Transaction }) => {
        if (Object.keys((transaction as unknown as { meta: object }).meta).some((key) => key.startsWith('wordTabLayout$'))) {
          editor.off('transaction', painted);resolve();
        }
      };
      editor.on('transaction', painted);
    });
  });
  // A real viewport resize changes the right-aligned field's margin constraint
  // and forces a spacer repaint while the native caret is ahead of the model.
  await page.setViewportSize({ width: 600, height: 1200 });
  await page.evaluate(() => (window as unknown as { tabRepaint: Promise<void> }).tabRepaint);
  await expect.poll(() => tabs.last().evaluate((tab) => tab.getBoundingClientRect().width)).not.toBe(initialWidth);
  await expect(tabs).toHaveCount(2);
  expect(await selection()).toEqual({ model: 1, native: 1 });
  await page.keyboard.up('ArrowRight');
  await page.evaluate(() => { (window as unknown as { holdTabSelection: boolean }).holdTabSelection = false; });
  await page.keyboard.type('X');await expect(story.locator('p')).toHaveText('LXeft\tCenter\tRight');
  await page.keyboard.press('Control+z');await expect(story.locator('p')).toHaveText('Left\tCenter\tRight');
});
