import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
for (const kind of ['nextPage', 'continuous', 'evenPage', 'oddPage'])
  test(`Fresh sections ${kind}: repeated insertion, stories and recovery`, async ({ page }) => {
    test.setTimeout(120000);
    const run = process.env.NOFFICE_SECTION_INSERT_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Invalid section insertion run');
    const root = `.local/word-section-insertion/fresh-${run}/${kind}`;
    await fs.mkdir(root, { recursive: true });
    await page.context().grantPermissions(['local-fonts']);
    await page.goto('/');
    await page.getByRole('button', { name: 'Start Word', exact: true }).click();
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    await editor.fill('Alpha.Beta');
    await page.keyboard.press('Control+a');
    await page.getByRole('combobox', { name: 'Font family', exact: true }).selectOption('Arial');
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: /^Section 1 — Default header/ }).click();
    const header = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    await header.fill('Shared header');
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    const model = () =>
      editor.evaluate((el) =>
        (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getJSON(),
      );
    const bodyText = () =>
      editor.evaluate((el) => {
        const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
        return e.state.doc.textBetween(0, e.state.doc.content.size, '\r') + '\r';
      });
    const caret = async (from: number, to = from) =>
      editor.evaluate(
        (el, range) => {
          const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
          e.commands.setTextSelection(range);
          e.view.focus();
        },
        { from, to },
      );
    const hashes: Record<string, string> = {},
      states: Record<string, unknown> = {};
    const capture = async (stage: string) => {
      await waitWordLayout(editor);
      states[stage] = await model();
      for (const [ext, label] of [
        ['docx', 'DOCX file Editable in Microsoft Word'],
        ['pdf', 'PDF file'],
      ]) {
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        const pending = page.waitForEvent('download');
        await page.getByRole('button', { name: label, exact: true }).click();
        const path = `${root}/${stage}.${ext}`;
        await (await pending).saveAs(path);
        hashes[`${stage}.${ext}`] = hash(await fs.readFile(path));
      }
    };
    await capture('source');
    await caret(4);
    const before = await model();
    const menu = page.getByRole('combobox', { name: 'Insert section break', exact: true });
    await menu.selectOption(kind);
    expect(await bodyText()).toBe('Alp\rha.Beta\r');
    const one = await model();
    await capture('inserted');
    await caret(6);
    await menu.selectOption('oddPage');
    expect(await bodyText()).toBe('Alp\r\rha.Beta\r');
    const two = await model();
    await capture('repeated');
    await editor.focus();
    await page.keyboard.press('Control+z');
    await expect.poll(model).toEqual(one);
    await page.keyboard.press('Control+z');
    await expect.poll(model).toEqual(before);
    await page.keyboard.press('Control+y');
    await page.keyboard.press('Control+y');
    await expect.poll(model).toEqual(two);
    await caret(4, 6);
    await page.keyboard.press('Delete');
    const deleted = await model();
    expect(deleted.attrs!.wordSectionState.breaks).toHaveLength(1);
    await page.keyboard.press('Control+z');
    await expect.poll(model).toEqual(two);
    await page.keyboard.press('Control+y');
    await expect.poll(model).toEqual(deleted);
    await capture('deleted');
    // A recoverable export failure must neither lose the sections nor dirty history.
    await page.context().clearPermissions();
    await page.context().grantPermissions([]);
    expect(
      await page.evaluate(
        async () =>
          (await navigator.permissions.query({ name: 'local-fonts' as PermissionName })).state,
      ),
    ).toBe('denied');
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    await page.getByRole('button', { name: 'PDF file', exact: true }).click();
    await expect(page.getByRole('alert')).toContainText('Font access was not granted');
    expect(await model()).toEqual(deleted);
    await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
    await page.context().grantPermissions(['local-fonts']);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const savedHtml = await editor.evaluate((el) =>
      (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getHTML(),
    );
    const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    await waitWordLayout(editor);
    // Fresh toolbar marks normalize null/empty CSS values and optional font
    // quotes on HTML reload; serialized formatting and section/story state must
    // still be exact, separately from full-model history above.
    expect(
      await editor.evaluate((el) =>
        (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getHTML(),
      ),
    ).toBe(savedHtml);
    expect((await model()).attrs).toEqual(deleted.attrs);
    await capture('reloaded');
    await page.locator('input[type=file][multiple]').setInputFiles(root + '/deleted.docx');
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'deleted',
    );
    await expect
      .poll(() =>
        editor.evaluate(
          (el) => !!(el as HTMLElement & { editor?: import('@tiptap/core').Editor }).editor,
        ),
      )
      .toBe(true);
    await waitWordLayout(editor);
    expect(await bodyText()).toBe('Alp\rha.Beta\r');
    await capture('reimported');
    expect(hashes['reimported.docx']).toBe(hashes['deleted.docx']);
    await fs.writeFile(
      root + '/report.json',
      JSON.stringify(
        {
          kind,
          hashes,
          states,
          buildHash: await wordStoryBuildHash(),
          testHash: hash(await fs.readFile('tests/word-section-insert-fresh.spec.ts')),
          scope:
            'Fresh file and header authoring, repeated section menu, actual Delete, atomic history, denied-font recovery, saved model reload, actual DOCX/PDF and reimport. Native file comparisons remain independent.',
        },
        null,
        2,
      ),
    );
  });
