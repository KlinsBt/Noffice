import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import {
  clickWordText,
  wordPlacements,
  wordLineOrigins,
  waitWordLayout,
} from './word-pagination-helpers';

const references = JSON.parse(
  readFileSync('tests/fixtures/native-word-inline-flow.json', 'utf8'),
) as typeof import('./fixtures/native-word-inline-flow.json');
for (const name of Object.keys(references.cases) as (keyof typeof references.cases)[])
  test(`Word inline flow ${name}: delete, restore, history, reload, caret and actual exports`, async ({
    page,
  }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    const root = `.local/word-inline-acceptance/${name}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(`tests/fixtures/word-inline-flow-${name}.docx`);
    const reference = references.cases[name];
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    expect(hash(source)).toBe(reference.sourceSha256);
    const expected = reference.native.paragraphs.map((p) =>
      p.characters.map((c) => ({ text: c.text, page: c.page })),
    );
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
    await upload(`Inline ${name}`, source);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    const check = async () => {
      await expect.poll(() => wordPlacements(editor)).toEqual(expected);
      await expect(page.locator('.section-page')).toHaveCount(
        Math.max(reference.native.paragraphs[0].caret.page, ...expected.flat().map((c) => c.page)),
      );
    };
    const caret = async () => {
      // Reimport can briefly invalidate surfaces after the character map has
      // matched. Measure the caret only after the actual view has settled.
      await waitWordLayout(editor);
      await editor.focus();
      await page.keyboard.press('Control+End');
      return editor.evaluate(async (root) => {
        const selection = window.getSelection()!;
        let r = selection.getRangeAt(0).getBoundingClientRect();
        // Chromium returns no rectangle for a collapsed range after an inline
        // atom. The real trailing BR is its empty insertion line; subsequent
        // typing is independently checked against Word's edited character map.
        if (!r.height) {
          const trailing = root.querySelector('p:last-child > br.ProseMirror-trailingBreak');
          if (!trailing) throw Error('Missing terminal insertion line');
          r = trailing.getBoundingClientRect();
        }
        const pages = [...document.querySelectorAll<HTMLElement>('.section-page')];
        const index = pages.findIndex((p) => {
          const b = p.getBoundingClientRect();
          return r.top + r.height / 2 >= b.top && r.top + r.height / 2 < b.bottom;
        });
        if (index < 0) {
          const frames = [];
          for (let frame = 0; frame < 4; frame++) {
            const selected = window.getSelection()!;
            frames.push({
              caret: selected.getRangeAt(0).getBoundingClientRect().toJSON(),
              anchor: { node: selected.anchorNode?.nodeName, offset: selected.anchorOffset },
              pages: [...document.querySelectorAll<HTMLElement>('.section-page')].map((page) =>
                page.getBoundingClientRect().toJSON(),
              ),
            });
            await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
          }
          throw Error(`Caret lies outside page surfaces: ${JSON.stringify(frames)}`);
        }
        const p = pages[index],
          b = p.getBoundingClientRect(),
          scale = b.width / parseFloat(getComputedStyle(p).width);
        return {
          page: index + 1,
          x: ((r.left - b.left) * 0.75) / scale,
          y: ((r.top - b.top) * 0.75) / scale,
        };
      });
    };
    const download = async (stage: string) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page
        .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
        .click();
      const bytes = await fs.readFile((await (await pending).path())!);
      await fs.writeFile(`${root}/${stage}.docx`, bytes);
      return bytes;
    };
    await check();
    const positions = await wordPlacements(editor),
      lines = await wordLineOrigins(editor),
      sourceCaret = await caret();
    expect(sourceCaret.page).toBe(reference.native.paragraphs[0].caret.page);
    expect(Math.abs(sourceCaret.x - reference.native.paragraphs[0].caret.x)).toBeLessThanOrEqual(
      0.15,
    );
    // Restore the imported inline structure through history. Native insertion
    // commands create paragraphs and have a separate native UI fixture matrix.
    const original = expected[0].map((c) => c.text).join(''),
      at = original.search(/[\f\u000e]/);
    for (let i = 0; i < 5; i++)
      await page.getByRole('button', { name: 'Zoom out', exact: true }).click();
    await clickWordText(page, editor, original.startsWith('before') ? 'before' : 'after');
    await page.keyboard.press('Control+Home');
    for (let i = 0; i < at; i++) await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Delete');
    const semantic = () =>
      editor.locator('[data-word-page-break],[data-word-column-break]').count();
    await expect.poll(semantic).toBe((original.match(/[\f\u000e]/g) || []).length - 1);
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await check();
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect.poll(semantic).toBe((original.match(/[\f\u000e]/g) || []).length - 1);
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await check();
    for (let i = 0; i < 5; i++)
      await page.getByRole('button', { name: 'Zoom in', exact: true }).click();
    // Persist the restored control before a separate typing history scenario.
    // Adjacent terminal break insertion and immediate typing may form one event.
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page
      .getByRole('button', { name: `Inline ${name}`, exact: true })
      .first()
      .click();
    await check();
    await caret();
    await page.keyboard.insertText('!');
    await expect(editor.locator('p')).toContainText('!');
    // Each fixture has room for one final character on Word's captured caret
    // line. Wait for that independent page expectation before saving geometry.
    await expect
      .poll(() => wordPlacements(editor))
      .toEqual(
        expected.map((characters, i) => [
          ...characters,
          { text: '!', page: reference.native.paragraphs[i].caret.page },
        ]),
      );
    const editedPositions = await wordPlacements(editor),
      editedLines = await wordLineOrigins(editor),
      editedCaret = await caret();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await check();
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page
      .getByRole('button', { name: `Inline ${name}`, exact: true })
      .first()
      .click();
    await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
    const edited = await download('browser-edited');
    // Closing/cancelling print must remove the temporary tree and leave editing intact.
    await page.evaluate(() => window.dispatchEvent(new Event('beforeprint')));
    await expect(page.locator('.word-fragment-print')).toHaveCount(1);
    await page.evaluate(() => window.dispatchEvent(new Event('afterprint')));
    await expect(page.locator('.word-fragment-print')).toHaveCount(0);
    await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
    const editedPDF = await page.pdf({
      path: `${root}/browser-edited-print.pdf`,
      preferCSSPageSize: true,
    });
    await expect(page.locator('.word-fragment-print')).toHaveCount(0);
    await upload(`Inline ${name} edited`, edited);
    await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
    await caret();
    await page.keyboard.press('Backspace');
    await check();
    const restored = await download('browser');
    const pdf = await page.pdf({ path: `${root}/browser-print.pdf`, preferCSSPageSize: true });
    await upload(`Inline ${name} restored`, restored);
    await check();
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify(
        {
          passed: true,
          sourceHash: hash(source),
          exportHash: hash(restored),
          editedHash: hash(edited),
          printHash: hash(pdf),
          editedPrintHash: hash(editedPDF),
          positions,
          lines,
          editedPositions,
          editedLines,
          caret: [sourceCaret],
          editedCaret: [editedCaret],
          browser: page.context().browser()!.version(),
        },
        null,
        2,
      ) + '\n',
    );
  });
