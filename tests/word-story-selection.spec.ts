import { test, expect } from '@playwright/test';

for (const kind of ['header', 'footer'])
  test(`Word ${kind} repeatedly cancels, reopens and replaces a native line selection`, async ({
    page,
  }, info) => {
    test.setTimeout(90000);
    await page.addInitScript(() => {
      const record: unknown[] = [];
      (window as unknown as { storyEvents: unknown[] }).storyEvents = record;
      for (const type of ['keydown', 'keyup', 'beforeinput', 'input', 'selectionchange'])
        document.addEventListener(
          type,
          (event) => {
            const element = document.querySelector('[aria-label="Header or footer text"]') as
              | (HTMLElement & {
                  editor?: {
                    state: {
                      selection: { anchor: number; head: number };
                      doc: { textContent: string };
                    };
                  };
                })
              | null;
            if (!element?.editor) return;
            const selection = element.editor.state.selection;
            record.push({
              type,
              key: (event as KeyboardEvent).key,
              data: (event as InputEvent).data,
              selection: { anchor: selection.anchor, head: selection.head },
              dom: document.getSelection()?.toString(),
              text: element.editor.state.doc.textContent,
              focused: document.activeElement === element,
            });
            if (record.length > 500) record.shift();
          },
          true,
        );
    });
    await page.goto('/');
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles('tests/fixtures/word-side-stories/default.docx');
    const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    const open = async () => {
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      await page.getByRole('button', { name: `Section 1 — Default ${kind}`, exact: true }).click();
      await expect(editor).toHaveText(`${kind === 'header' ? 'Header' : 'Footer'} default`);
      await editor.focus();
    };
    try {
      for (let iteration = 0; iteration < 20; iteration++) {
        await open();
        await page.keyboard.press('Control+End');
        await page.keyboard.insertText(' cancelled');
        await page.getByRole('button', { name: 'Cancel', exact: true }).click();
        await open();
        await page.keyboard.press('Control+Home');
        await page.keyboard.press('Shift+End');
        await page.keyboard.insertText('Replacement');
        await expect(editor, `replacement ${iteration}`).toHaveText('Replacement');
        await page.getByRole('button', { name: 'Cancel', exact: true }).click();
      }
    } catch (error) {
      await info.attach('selection-events', {
        body: JSON.stringify(
          await page.evaluate(() => (window as unknown as { storyEvents: unknown[] }).storyEvents),
          null,
          2,
        ),
        contentType: 'application/json',
      });
      throw error;
    }
  });
