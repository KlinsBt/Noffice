import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import reference from './fixtures/word-section-insertion/reference.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
for (const sample of reference.rows)
  test(`Section menu ${sample.name}: editing, history, reload and files`, async ({ page }) => {
    test.setTimeout(90000);
    const run = process.env.NOFFICE_SECTION_INSERT_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Invalid section insertion run');
    const root = `.local/word-section-insertion/${run}/${sample.name}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(sample.source);
    expect(hash(source)).toBe(sample.sourceHash);
    await page.context().grantPermissions(['local-fonts']);
    await page.goto('/');
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
    const verify = async (expected: string) => {
      const state = await editor.evaluate((el) => {
        const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
        const live = e.state.doc.attrs.wordSectionState;
        const boundaries = new Set(live.breaks.map((b: { paragraph: number }) => b.paragraph));
        let text = '',
          ordinal = 0;
        e.state.doc.descendants((node) => {
          if (!node.isTextblock) return;
          text += node.textContent + (boundaries.has(ordinal++) ? '\f' : '\r');
          return false;
        });
        return {
          text,
          sections: live.breaks.length + 1,
          from: e.state.selection.from,
          to: e.state.selection.to,
        };
      });
      expect(state.text).toBe(expected);
      expect(state.sections).toBe(sample.sections.length);
      return state;
    };
    const from = await editor.evaluate((el, selection) => {
      const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
      const paragraphs: { from: number; length: number }[] = [];
      e.state.doc.descendants((node, from) => {
        if (!node.isTextblock) return;
        paragraphs.push({ from, length: node.textContent.length });
        return false;
      });
      const position = (offset: number) => {
        let n = 0;
        for (const p of paragraphs) {
          if (offset <= n + p.length) return p.from + 1 + offset - n;
          n += p.length + 1;
        }
        throw Error('Native selection outside model');
      };
      const from = position(selection.from);
      e.commands.setTextSelection({ from, to: position(selection.to) });
      return from;
    }, sample.selection);
    const before = await model();
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page
      .getByRole('combobox', { name: 'Insert section break', exact: true })
      .selectOption(sample.kind);
    await waitWordLayout(editor);
    const state = await verify(sample.after);
    expect(state.from).toBe(from + 2);
    expect(state.to).toBe(from + 2);
    const after = await model();
    await editor.focus();
    await page.keyboard.press('Control+z');
    await expect.poll(model).toEqual(before);
    await page.keyboard.press('Control+y');
    await expect.poll(model).toEqual(after);
    // Actual typing is a separate undo event and preserves the inserted section.
    await page.keyboard.insertText('X');
    const typedText =
      sample.after.slice(0, sample.afterSelection.from) +
      'X' +
      sample.after.slice(sample.afterSelection.from);
    await verify(typedText);
    await page.keyboard.press('Control+z');
    await expect.poll(model).toEqual(after);
    await waitWordLayout(editor);
    const hashes: Record<string, string> = {},
      pages: Record<string, number> = {};
    const capture = async (stage: string) => {
      await waitWordLayout(editor);
      pages[stage] = await page.locator('.section-page').count();
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
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await capture('edited');
    const sourceZip = await JSZip.loadAsync(source),
      output = await JSZip.loadAsync(await fs.readFile(root + '/edited.docx'));
    for (const path of Object.keys(sourceZip.files).filter((p) =>
      /^word\/(header|footer).*\.xml$/.test(p),
    ))
      expect(await output.file(path)!.async('uint8array')).toEqual(
        await sourceZip.file(path)!.async('uint8array'),
      );
    const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    await waitWordLayout(editor);
    expect(await model()).toEqual(after);
    await verify(sample.after);
    await capture('reloaded');
    await page.locator('input[type=file][multiple]').setInputFiles(root + '/edited.docx');
    await waitWordLayout(editor);
    await verify(sample.after);
    await capture('reimported');
    expect(hashes['reimported.docx']).toBe(hashes['edited.docx']);
    const receipt = {
      name: sample.name,
      sourceHash: sample.sourceHash,
      before,
      after,
      hashes,
      pages,
      expectedPages: sample.pages,
      buildHash: await wordStoryBuildHash(),
      testHash: hash(await fs.readFile('tests/word-section-insert.spec.ts')),
      referenceHash: hash(
        await fs.readFile('tests/fixtures/word-section-insertion/reference.json'),
      ),
      scope:
        'Actual section menu, model-positioned selection, real typing and history, exact saved model, three DOCX/PDF stages, retained story bytes. Native reopening and PDF geometry are independent checks.',
    };
    await fs.writeFile(root + '/report.json', JSON.stringify(receipt, null, 2));
    expect(pages).toEqual({
      edited: sample.pages,
      reloaded: sample.pages,
      reimported: sample.pages,
    });
  });
