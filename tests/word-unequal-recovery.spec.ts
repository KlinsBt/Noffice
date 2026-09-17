import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { clickWordText, wordPlacements, waitWordLayout } from './word-pagination-helpers';

for (const order of ['narrow', 'wide'])
  test(`Word unequal ${order} columns preserve wrapped edits and recover pagination`, async ({
    page,
  }) => {
    const name = `${order}-first-hard`,
      root = `.local/word-unequal-recovery/${name}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(`tests/fixtures/word-unequal-flow-${name}.docx`);
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    const upload = async (label: string, bytes: Buffer) => {
      await page.locator('input[type=file][multiple]').setInputFiles({
        name: `${label}.docx`,
        buffer: bytes,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
        label,
      );
    };
    await page.setViewportSize({ width: 1500, height: 1200 });
    await page.goto('/');
    await upload(`Recovery ${name}`, source);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(page.locator('.section-page')).toHaveCount(2);
    await waitWordLayout(editor);
    const original = await wordPlacements(editor),
      text = await editor.textContent();
    await clickWordText(page, editor, 'u26');
    await page.keyboard.press('Control+End');
    const suffix = ' ' + 'oversizedword'.repeat(20);
    await page.keyboard.insertText(suffix);
    const fallback = async () => {
      await expect(
        page.getByText('Continuous view: this content needs paragraph pagination.', {
          exact: true,
        }),
      ).toBeVisible();
      await expect(page.locator('.section-page')).toHaveCount(0);
      await expect(editor).toHaveText(text! + suffix);
    };
    await fallback();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect.poll(() => wordPlacements(editor)).toEqual(original);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await fallback();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page
      .getByRole('button', { name: `Recovery ${name}`, exact: true })
      .first()
      .click();
    await fallback();
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const output = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/browser-wrapped.docx`, output);
    const before = await JSZip.loadAsync(source),
      after = await JSZip.loadAsync(output);
    const sourceXml = await before.file('word/document.xml')!.async('string');
    const outputXml = await after.file('word/document.xml')!.async('string');
    expect(outputXml.match(/<w:cols\b[\s\S]*?<\/w:cols>/)?.[0]).toBe(
      sourceXml.match(/<w:cols\b[\s\S]*?<\/w:cols>/)?.[0],
    );
    await upload(`Reimport ${name}`, output);
    await fallback();
    await editor.click();
    await page.keyboard.press('Control+End');
    await page.keyboard.down('Shift');
    for (let i = 0; i < suffix.length; i++) await page.keyboard.press('ArrowLeft');
    await page.keyboard.up('Shift');
    await page.keyboard.press('Backspace');
    await expect.poll(() => wordPlacements(editor)).toEqual(original);
    await expect(editor).toHaveText(text!);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify(
        {
          passed: true,
          sourceHash: hash(source),
          exportHash: hash(output),
          suffix,
          browser: page.context().browser()!.version(),
          scope:
            'Guarded oversized words, history, persisted edits, retained columns, actual DOCX reimport and keyboard recovery. Character-level wrapping remains open.',
        },
        null,
        2,
      ) + '\n',
    );
  });
