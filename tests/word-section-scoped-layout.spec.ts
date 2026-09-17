import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/word-section-scoped-layout/reference.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
const operations = (name: string): { label: string; value: string }[] => {
  switch (name) {
    case 'margin-normal':
      return [{ label: 'Margins', value: 'normal' }];
    case 'margin-wide':
      return [{ label: 'Margins', value: 'wide' }];
    case 'a4':
      return [{ label: 'Paper size', value: 'a4' }];
    case 'start-next':
      return [{ label: 'Section start', value: 'nextPage' }];
    case 'margin-narrow':
      return [{ label: 'Margins', value: 'narrow' }];
    case 'landscape':
      return [{ label: 'Page orientation', value: 'landscape' }];
    case 'portrait':
      return [{ label: 'Page orientation', value: 'portrait' }];
    case 'letter':
      return [{ label: 'Paper size', value: 'letter' }];
    case 'start-even':
      return [{ label: 'Section start', value: 'evenPage' }];
    case 'start-odd':
      return [{ label: 'Section start', value: 'oddPage' }];
    case 'start-continuous':
      return [{ label: 'Section start', value: 'continuous' }];
    case 'landscape-letter':
      return [...operations('landscape'), ...operations('letter')];
    case 'letter-landscape':
      return [...operations('letter'), ...operations('landscape')];
    default:
      throw Error('Unknown native operation');
  }
};
for (const sample of reference.rows)
  test(`Scoped page setup ${sample.name}: history, reload and actual files`, async ({ page }) => {
    test.setTimeout(90000);
    const run = process.env.NOFFICE_SCOPED_SECTION_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Use versioned scoped-section evidence.');
    const root = `.local/word-section-scoped-layout/${run}/${sample.name}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(sample.fixture);
    expect(hash(source)).toBe(sample.sourceHash);
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
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
    const state = () =>
      editor.evaluate((el) => {
        const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
        return {
          html: e.getHTML(),
          sections: e.state.doc.attrs.wordSectionState,
          stories: e.state.doc.attrs.wordStories,
          text: e.state.doc.textBetween(0, e.state.doc.content.size, '\r') + '\r',
        };
      });
    await editor.evaluate((el, target) => {
      const e = (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor;
      const live = e.state.doc.attrs.wordSectionState;
      const ordinal = target <= 1 ? 0 : live.breaks[target - 2].paragraph + 1;
      let index = 0;
      e.state.doc.descendants((node, from) => {
        if (!node.isTextblock) return;
        if (index++ === ordinal) e.commands.setTextSelection(from + 1);
        return false;
      });
    }, sample.target);
    const before = await model();
    let changed = 0;
    await page.getByRole('button', { name: 'Layout', exact: true }).click();
    await page
      .getByRole('combobox', { name: 'Apply page setup to', exact: true })
      .selectOption(sample.target === 0 ? 'document' : 'section');
    for (const operation of operations(sample.operation)) {
      const before = await model();
      await page
        .getByRole('combobox', { name: operation.label, exact: true })
        .selectOption(operation.value);
      await waitWordLayout(editor);
      if (JSON.stringify(await model()) !== JSON.stringify(before)) changed++;
    }
    expect(changed === 0).toBe(sample.native.unchanged);
    const after = await model(),
      saved = await state();
    expect(saved.text).toBe(sample.native.edited.text.replaceAll('\f', '\r'));
    await editor.focus();
    for (let i = 0; i < changed; i++) await page.keyboard.press('Control+z');
    await expect.poll(model).toEqual(before);
    for (let i = 0; i < changed; i++) await page.keyboard.press('Control+y');
    await expect.poll(model).toEqual(after);
    // A following text edit has its own history event.
    await page.keyboard.insertText('X');
    await expect.poll(async () => (await state()).text).not.toEqual(saved.text);
    await page.keyboard.press('Control+z');
    await expect.poll(model).toEqual(after);
    const outputs: Record<string, string> = {},
      pages: Record<string, number> = {};
    const capture = async (stage: string) => {
      await waitWordLayout(editor);
      pages[stage] = await page.locator('.section-page').count();
      for (const [extension, label] of [
        ['docx', 'DOCX file Editable in Microsoft Word'],
        ['pdf', 'PDF file'],
      ]) {
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        const pending = page.waitForEvent('download');
        await page.getByRole('button', { name: label, exact: true }).click();
        const path = `${root}/${stage}.${extension}`;
        await (await pending).saveAs(path);
        outputs[`${stage}.${extension}`] = hash(await fs.readFile(path));
      }
    };
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await capture('edited');
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: sample.name, exact: true }).click();
    await waitWordLayout(editor);
    expect(await state()).toEqual(saved);
    await capture('reloaded');
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/edited.docx`);
    await waitWordLayout(editor);
    expect((await state()).text).toBe(saved.text);
    await capture('reimported');
    expect(outputs['reimported.docx']).toBe(outputs['edited.docx']);
    await fs.writeFile(
      `${root}/report.json`,
      JSON.stringify(
        {
          name: sample.name,
          sourceHash: sample.sourceHash,
          outputs,
          pages,
          before,
          after,
          errors,
          buildHash: await wordStoryBuildHash(),
          testHash: hash(await fs.readFile('tests/word-section-scoped-layout.spec.ts')),
          referenceHash: hash(
            await fs.readFile('tests/fixtures/word-section-scoped-layout/reference.json'),
          ),
          scope:
            'Current-section page setup controls; actual typing, atomic metadata history, no-op behavior, saved state, three actual DOCX/PDF stages and unchanged reimport download. Native geometry and PDF comparisons are separate.',
        },
        null,
        2,
      ),
    );
    expect(errors).toEqual([]);
    expect(pages).toEqual({
      edited: sample.native.edited.pages,
      reloaded: sample.native.edited.pages,
      reimported: sample.native.edited.pages,
    });
  });
