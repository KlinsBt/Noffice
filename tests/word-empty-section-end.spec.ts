import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/word-empty-section-end/reference.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
for (const sample of reference.rows)
  test(`Empty section end ${sample.name}: typing, history, reload and files`, async ({ page }) => {
    test.setTimeout(90000);
    const run = process.env.NOFFICE_EMPTY_SECTION_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Invalid empty section evidence run');
    const root = `.local/word-empty-section-end/${run}/${sample.name}`;
    await fs.mkdir(root, { recursive: true });
    await page.context().grantPermissions(['local-fonts']);
    await page.goto('/');
    const source = await fs.readFile(sample.source);
    expect(hash(source)).toBe(sample.sourceHash);
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: sample.name + '.docx',
      buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    await waitWordLayout(editor);
    const model = () =>
      editor.evaluate((el) =>
        (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getJSON(),
      );
    const text = () =>
      editor.evaluate((el) => {
        const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
        return e.state.doc.textBetween(0, e.state.doc.content.size, '\r') + '\r';
      });
    const hashes: Record<string, string> = {};
    const capture = async (stage: string, docx = true) => {
      await waitWordLayout(editor);
      for (const [ext, label] of [
        ['docx', 'DOCX file Editable in Microsoft Word'],
        ['pdf', 'PDF file'],
      ]) {
        if (ext === 'docx' && !docx) continue;
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        const waiting = page.waitForEvent('download');
        await page.getByRole('button', { name: label, exact: true }).click();
        const path = `${root}/${stage}.${ext}`;
        await (await waiting).saveAs(path);
        hashes[stage + '.' + ext] = hash(await fs.readFile(path));
      }
    };
    await capture('source');
    expect(hashes['source.docx']).toBe(sample.sourceHash);
    const before = await model();
    await editor.evaluate((el) => {
      const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
      let from = 0;
      for (let i = 0; i < e.state.doc.childCount - 2; i++) from += e.state.doc.child(i).nodeSize;
      e.commands.setTextSelection(from + 1);
      e.view.focus();
    });
    await page.keyboard.insertText('X');
    await waitWordLayout(editor);
    expect(await text()).toBe(sample.typedSnapshot.text.replaceAll('\f', '\r'));
    const edited = await model();
    await capture('edited');
    await editor.focus();
    await page.keyboard.press('Control+z');
    await expect.poll(model).toEqual(before);
    await capture('restored', false);
    await editor.focus();
    await page.keyboard.press('Control+y');
    await expect.poll(model).toEqual(edited);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const savedHtml = await editor.evaluate((el) =>
      (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getHTML(),
    );
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: sample.name, exact: true }).click();
    await waitWordLayout(editor);
    expect(
      await editor.evaluate((el) =>
        (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getHTML(),
      ),
    ).toBe(savedHtml);
    expect((await model()).attrs).toEqual(edited.attrs);
    await capture('reloaded');
    await page.locator('input[type=file][multiple]').setInputFiles(root + '/edited.docx');
    await waitWordLayout(editor);
    expect(await text()).toBe(sample.typedSnapshot.text.replaceAll('\f', '\r'));
    await capture('reimported');
    expect(hashes['reimported.docx']).toBe(hashes['edited.docx']);
    await fs.writeFile(
      root + '/report.json',
      JSON.stringify(
        {
          name: sample.name,
          sourceHash: sample.sourceHash,
          hashes,
          before,
          edited,
          buildHash: await wordStoryBuildHash(),
          testHash: hash(await fs.readFile('tests/word-empty-section-end.spec.ts')),
          referenceHash: hash(
            await fs.readFile('tests/fixtures/word-empty-section-end/reference.json'),
          ),
          scope:
            'Native empty-section source, real X entry and exact full-model history, exact serialized formatting/section metadata after reload, original-byte preservation, actual DOCX/PDF and reimport. Geometry and native reopen are checked independently.',
        },
        null,
        2,
      ),
    );
  });
