import { test, expect } from '@playwright/test';

for (const kind of ['body', 'header']) test(`${kind} line repaint preserves native navigation before selectionchange delivery`, async ({ page }) => {
  await page.addInitScript(() => {
    const probe = window as unknown as { focusSettled?: Promise<void>; holdLineSelection?: boolean;
      holdLineFrames?: boolean; releaseLineFrames?: () => void };
    const requestFrame = window.requestAnimationFrame.bind(window), queued: FrameRequestCallback[] = [];
    window.requestAnimationFrame = callback => requestFrame(time => {
      if (probe.holdLineFrames) queued.push(callback); else callback(time);
    });
    probe.releaseLineFrames = () => {
      probe.holdLineFrames = false;
      for (const callback of queued.splice(0)) requestFrame(callback);
    };
    const schedule = window.setTimeout.bind(window);
    window.setTimeout = ((handler: TimerHandler, delay?: number, ...args: unknown[]) => {
      if (typeof handler === 'function' && String(handler).includes('domSelectionRange') &&
        String(handler).includes('currentSelection') && ['Document text', 'Header or footer text'].includes(document.activeElement?.getAttribute('aria-label') || '')) {
        let finish!: () => void; probe.focusSettled = new Promise<void>(resolve => { finish = resolve; });
        return schedule(() => { try { handler(...args); } finally { finish(); } }, delay);
      }
      return schedule(handler, delay, ...args);
    }) as typeof window.setTimeout;
    document.addEventListener('selectionchange', event => {
      if (probe.holdLineSelection && document.getSelection()?.focusNode?.parentElement?.closest('[aria-label="Document text"],[aria-label="Header or footer text"]'))
        event.stopImmediatePropagation();
    }, true);
  });
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(`tests/fixtures/word-script-font-${kind === 'body' ? 'body' : 'headers'}-20-subscript.docx`);
  if (kind === 'header') {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
  }
  const editor = page.getByRole('textbox', { name: kind === 'body' ? 'Document text' : 'Header or footer text', exact: true });
  if (kind === 'header') {
    // A plain story isolates line paint from both page and tab decorators.
    await editor.focus(); await page.keyboard.press('Control+a'); await page.keyboard.insertText('AB');
    await expect(editor.locator('p').first()).toHaveText('AB');
    // Commit the plain-text setup and reopen its editor so the later typing
    // assertion has its own history, independent of the fixture-preparation edit.
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
  }
  await expect(editor.locator('p').first()).toHaveAttribute('data-word-measured-leading', 'true');
  await editor.focus();
  await page.waitForFunction(() => !!(window as unknown as { focusSettled?: Promise<void> }).focusSettled);
  await page.evaluate(() => (window as unknown as { focusSettled: Promise<void> }).focusSettled);
  await page.keyboard.press('Control+Home');
  await page.evaluate(() => {
    const probe = window as unknown as { holdLineSelection: boolean; holdLineFrames: boolean };
    probe.holdLineSelection = true; probe.holdLineFrames = true;
  });
  await page.keyboard.down('ArrowRight');
  const selection = () => editor.evaluate(host => {
    const e = (host as HTMLElement & { editor: import('@tiptap/core').Editor }).editor, dom = document.getSelection()!;
    return { model: e.state.selection.$head.parentOffset,
      native: e.state.doc.resolve(e.view.posAtDOM(dom.focusNode!, dom.focusOffset)).parentOffset };
  });
  expect(await selection()).toEqual({ model: 0, native: 1 });
  await editor.evaluate(host => {
    const e = (host as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    (window as unknown as { lineRepaint: Promise<void> }).lineRepaint = new Promise(resolve => {
      const painted = ({ transaction }: { transaction: import('@tiptap/pm/state').Transaction }) => {
        if (Object.keys((transaction as unknown as { meta: object }).meta).some(key => key.startsWith('wordFontLineMetrics$'))) {
          e.off('transaction', painted); resolve();
        }
      };
      e.on('transaction', painted);
    });
    // Change the same CSS zoom consumed by the real view control, while keeping
    // the native arrow held so keyup cannot synchronize the caret first.
    ((host.closest('.paper-wrap') || host) as HTMLElement).style.zoom = '.9';
    host.dispatchEvent(new Event('wordlayoutchange'));
    (window as unknown as { releaseLineFrames: () => void }).releaseLineFrames();
  });
  await page.evaluate(() => (window as unknown as { lineRepaint: Promise<void> }).lineRepaint);
  expect(await selection()).toEqual({ model: 1, native: 1 });
  await page.keyboard.up('ArrowRight');
  await page.evaluate(() => { (window as unknown as { holdLineSelection: boolean }).holdLineSelection = false; });
  await page.keyboard.insertText('X'); await expect(editor.locator('p').first()).toHaveText(kind === 'body' ? 'AX\tB' : 'AXB');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText(kind === 'body' ? 'A\tB' : 'AB');
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText(kind === 'body' ? 'AX\tB' : 'AXB');
});
