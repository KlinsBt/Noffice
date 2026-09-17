import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
test('Fresh scoped page setup: insertion inheritance, global controls and exports', async ({
  page,
}) => {
  test.setTimeout(120000);
  const run = process.env.NOFFICE_SCOPED_SECTION_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid section insertion run');
  const root = `.local/word-section-scoped-layout/fresh-${run}/authored`;
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
  await page.getByRole('button', { name: 'Layout', exact: true }).click();
  await page
    .getByRole('combobox', { name: 'Apply page setup to', exact: true })
    .selectOption('section');
  await page
    .getByRole('combobox', { name: 'Page orientation', exact: true })
    .selectOption('landscape');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page
    .getByRole('combobox', { name: 'Insert section break', exact: true })
    .selectOption('nextPage');
  expect(await bodyText()).toBe('Alp\rha.Beta\r');
  await page.getByRole('button', { name: 'Layout', exact: true }).click();
  await page
    .getByRole('combobox', { name: 'Apply page setup to', exact: true })
    .selectOption('section');
  await expect(page.getByRole('combobox', { name: 'Page orientation', exact: true })).toHaveValue(
    'landscape',
  );
  await page.getByRole('combobox', { name: 'Paper size', exact: true }).selectOption('letter');
  await expect(page.getByRole('combobox', { name: 'Page orientation', exact: true })).toHaveValue(
    'portrait',
  );
  await page.getByRole('combobox', { name: 'Margins', exact: true }).selectOption('narrow');
  const scoped = await model();
  await capture('scoped');
  await page
    .getByRole('combobox', { name: 'Apply page setup to', exact: true })
    .selectOption('document');
  await page
    .getByRole('combobox', { name: 'Page orientation', exact: true })
    .selectOption('portrait');
  const global = await model();
  await editor.focus();
  await page.keyboard.press('Control+z');
  await expect.poll(model).toEqual(scoped);
  await page.keyboard.press('Control+y');
  await expect.poll(model).toEqual(global);
  await page.keyboard.insertText('X');
  await expect.poll(bodyText).toBe('Alp\rXha.Beta\r');
  await page.keyboard.press('Control+z');
  await expect.poll(model).toEqual(global);
  await capture('global');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const saved = await editor.evaluate((el) => {
    const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
    return {
      html: e.getHTML(),
      sections: e.state.doc.attrs.wordSectionState,
      stories: e.state.doc.attrs.wordStories,
    };
  });
  const title = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: title, exact: true }).click();
  await waitWordLayout(editor);
  expect(
    await editor.evaluate((el) => {
      const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
      return {
        html: e.getHTML(),
        sections: e.state.doc.attrs.wordSectionState,
        stories: e.state.doc.attrs.wordStories,
      };
    }),
  ).toEqual(saved);
  await capture('reloaded');
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(root + '/global.docx');
  await waitWordLayout(editor);
  expect(await bodyText()).toBe('Alp\rha.Beta\r');
  await capture('reimported');
  expect(hashes['reimported.docx']).toBe(hashes['global.docx']);
  await fs.writeFile(
    root + '/report.json',
    JSON.stringify(
      {
        hashes,
        states,
        buildHash: await wordStoryBuildHash(),
        testHash: hash(await fs.readFile('tests/word-section-scoped-fresh.spec.ts')),
        scope:
          'Fresh scoped geometry before insertion is inherited; second-section letter/margins and whole-document orientation retain distinct sizes. Real history/typing/reload and all five DOCX/PDF stages are retained for independent native comparison.',
      },
      null,
      2,
    ),
  );
});
