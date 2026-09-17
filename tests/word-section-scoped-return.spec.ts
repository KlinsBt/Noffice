import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/word-section-scoped-layout/reference.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

test.skip(
  process.env.NOFFICE_SCOPED_NATIVE_RETURN !== '1',
  'Requires actual further Word edits from the native file verifier.',
);
const sourceRun = process.env.NOFFICE_SCOPED_NATIVE_SOURCE || 'exports-native-v1';
const run = process.env.NOFFICE_SCOPED_RETURN_RUN || 'returns-browser-v1';
const freshSource = process.env.NOFFICE_SCOPED_FRESH_NATIVE_SOURCE || 'fresh-exports-native-v1';
if (!/^exports-native-v\d+$/.test(sourceRun) || !/^returns-browser-v\d+$/.test(run))
  throw Error('Use versioned native return evidence.');
if (!/^fresh-exports-native-v\d+$/.test(freshSource)) throw Error('Invalid fresh native evidence.');
const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');

const cases = [
  ...reference.rows.map((row) => ({
    name: row.name,
    after: row.native.edited.text,
    sections: row.native.edited.sections,
    sourceRun,
    sourceBase: '.local/word-section-scoped-layout',
    folderRun: run,
    stage: 'edited-return',
  })),
  {
    name: 'authored',
    after: 'Alp\rha.Beta\r',
    sections: [{}, {}],
    sourceRun: freshSource,
    sourceBase: '.local/word-section-scoped-layout',
    folderRun: 'fresh-' + run,
    stage: 'global-return',
  },
];
for (const sample of cases)
  test(`Scoped section ${sample.name}: native return, history, reload and preserved original`, async ({
    page,
  }) => {
    const base = '.local/word-section-scoped-layout',
      sourceRoot = `${sample.sourceBase}/${sample.sourceRun}`;
    const receiptBytes = await fs.readFile(`${sourceRoot}/report.json`);
    const receipt = JSON.parse(receiptBytes.toString('utf8').replace(/^\uFEFF/, ''));
    const state = receipt.rows
      .find((row: { name: string }) => row.name === sample.name)
      .states.find((state: { stage: string }) => state.stage === sample.stage);
    const sourcePath = `${sourceRoot}/${sample.name}/${sample.stage}.docx`,
      source = await fs.readFile(sourcePath);
    expect(hash(source)).toBe(state.savedHash);
    const root = `${base}/${sample.folderRun}/${sample.name}`,
      outputs: Record<string, string> = {},
      errors: string[] = [];
    await fs.mkdir(root, { recursive: true });
    page.on('pageerror', (error) => errors.push(error.message));
    await page.context().grantPermissions(['local-fonts']);
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles(sourcePath);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    await waitWordLayout(editor);
    const text = () =>
      editor.evaluate((el) => {
        const editor = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
        return editor.state.doc.textBetween(0, editor.state.doc.content.size, '\r') + '\r';
      });
    const original = 'X' + sample.after.replaceAll('\f', '\r'),
      typed = original.slice(0, -1) + 'X\r';
    expect(
      await editor.evaluate(
        (el) =>
          (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.state.doc.attrs
            .wordSectionState.breaks.length + 1,
      ),
    ).toBe(sample.sections.length);
    expect(await text()).toEqual(original);
    await editor.focus();
    await page.keyboard.press('Control+End');
    await page.keyboard.type('X');
    await expect.poll(text).toEqual(typed);
    await page.keyboard.press('Control+z');
    await expect.poll(text).toEqual(original);
    await page.keyboard.press('Control+y');
    await expect.poll(text).toEqual(typed);
    await page.keyboard.press('Control+z');
    await expect.poll(text).toEqual(original);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const title = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: title, exact: true }).click();
    await waitWordLayout(editor);
    expect(await text()).toEqual(original);
    for (const [extension, label] of [
      ['docx', 'DOCX file Editable in Microsoft Word'],
      ['pdf', 'PDF file'],
    ]) {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: label, exact: true }).click();
      const path = `${root}/reloaded.${extension}`;
      await (await pending).saveAs(path);
      outputs[`reloaded.${extension}`] = hash(await fs.readFile(path));
    }
    expect(outputs['reloaded.docx']).toBe(hash(source));
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/report.json`,
      JSON.stringify(
        {
          name: sample.name,
          sourcePath,
          sourceHash: hash(source),
          outputs,
          errors,
          nativeReceiptHash: hash(receiptBytes),
          buildHash: await wordStoryBuildHash(),
          testHash: hash(await fs.readFile('tests/word-section-scoped-return.spec.ts')),
          scope:
            'Further Word edit returns to Noffice; real typing, undo/redo/undo and reload preserve its text and exact original DOCX bytes. Actual PDF retained for separate native comparison.',
        },
        null,
        2,
      ),
    );
  });
