import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/word-soft-hyphens/reference.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
for (const sample of reference.rows.filter(row => row.encoding === 'element'))
  test(`Insert optional hyphen using the keyboard: ${sample.name}`, async ({ page }) => {
    test.setTimeout(90000);
    const run = process.env.NOFFICE_HYPHEN_INPUT_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Invalid input evidence run');
    const root = `.local/word-soft-hyphen-input/${run}/${sample.name}`;
    await fs.mkdir(root, { recursive: true });
    await page.context().grantPermissions(['local-fonts']); await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles(sample.path);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true }); await waitWordLayout(editor);
    const model = () => editor.evaluate(el => (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getJSON());
    const text = () => editor.evaluate(el => {
      const editor = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
      const rows: string[] = []; editor.state.doc.forEach(p => rows.push(p.textBetween(0, p.content.size, '', n => {
        if (n.type.name !== 'wordHyphen' || n.attrs.kind !== 'optional') throw Error('Unexpected inline content');
        return '\u001f';
      }) + '\r')); return rows.join('');
    });
    const sourceModel = await model(), expected = sample.native.sourceText.replace('antidisestab', 'anti\u001fdisestab');
    expect(await text()).toBe(sample.native.sourceText);
    // The caret is placed after "anti" in paragraph2; the character itself
    // must be created by the actual browser keyboard shortcut.
    await editor.evaluate(el => {
      const editor = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
      editor.commands.setTextSelection(editor.state.doc.child(0).nodeSize + 5);
    });
    await editor.focus(); await page.keyboard.press('Control+-'); await waitWordLayout(editor);
    await expect.poll(text).toBe(expected); const insertedModel = await model();
    expect((await text()).split('\u001f')).toHaveLength(24);
    await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(sourceModel);
    await page.keyboard.press('Control+y'); await expect.poll(model).toEqual(insertedModel); await waitWordLayout(editor);
    const hashes: Record<string, string> = {};
    const capture = async (stage: string) => {
      for (const [extension, label] of [['docx', 'DOCX file Editable in Microsoft Word'], ['pdf', 'PDF file']]) {
        await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
        await page.getByRole('button', { name: label, exact: true }).click();
        const path = `${root}/${stage}.${extension}`; await (await pending).saveAs(path); hashes[stage + '.' + extension] = hash(await fs.readFile(path));
      }
    };
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible(); await capture('inserted');
    const title = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: title, exact: true }).click(); await waitWordLayout(editor);
    expect(await text()).toBe(expected); await capture('reloaded');
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/inserted.docx`); await waitWordLayout(editor);
    expect(await text()).toBe(expected); await capture('reimported'); expect(hashes['reimported.docx']).toBe(hashes['inserted.docx']);
    await fs.writeFile(`${root}/report.json`, JSON.stringify({ name: sample.name, sourceHash: sample.sourceHash,
      expected, sourceModel, insertedModel, hashes, buildHash: await wordStoryBuildHash(),
      testHash: hash(await fs.readFile('tests/word-hyphen-input.spec.ts')),
      scope: 'Actual browser Ctrl-minus insertion after a specified caret placement, full model undo/redo, reload/reimport and actual DOCX/PDF downloads. Native character insertion and native physical-keyboard evidence remain separate.' }, null, 2));
  });
