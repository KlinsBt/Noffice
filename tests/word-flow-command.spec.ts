import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { clickWordText, wordPlacements, wordLineOrigins } from './word-pagination-helpers';

const references = JSON.parse(
  readFileSync('tests/fixtures/native-word-ui-flow.json', 'utf8'),
) as typeof import('./fixtures/native-word-ui-flow.json');
for (const [name, reference] of Object.entries(references.cases))
  test(`Word native break command ${name}: paragraphs, history, reload and actual files`, async ({
    page,
  }) => {
    const root = `.local/word-ui-acceptance/${name}`;
    await fs.mkdir(root, { recursive: true });
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    const source = await fs.readFile(`tests/fixtures/${reference.inputFile}`);
    expect(hash(source)).toBe(reference.inputSha256);
    const zoomPrint = /^(page|column)-midline(-auto|-atLeast)?$/.test(name);
    const printRoute = ['column-midline', 'column-single-auto'].includes(name)
      ? 'menu'
      : zoomPrint || ['consecutive-column-atLeast'].includes(name)
        ? 'keyboard'
        : null;
    if (printRoute)
      await page.addInitScript(() => {
        // Headless Chromium's dialog is unavailable. Capture the application's
        // native print invocation and its before/after lifecycle; the workflow
        // below also writes real PDFs through Chromium's print engine.
        window.print = () => {
          const state = window as Window & { printInvocations?: number };
          state.printInvocations = (state.printInvocations || 0) + 1;
          window.dispatchEvent(new Event('beforeprint'));
          const tree = document.querySelector('.word-fragment-print');
          (window as Window & { printSnapshot?: unknown }).printSnapshot = {
            text: tree?.textContent,
            pages: tree?.querySelectorAll('.word-fragment-print-page').length,
          };
          window.dispatchEvent(new Event('afterprint'));
        };
      });
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
    await upload(`Command ${name}`, source);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    const positionsFor = (snapshot: typeof reference.native) =>
      snapshot.paragraphs.map((p) => p.characters.map((c) => ({ text: c.text, page: c.page })));
    const check = async (snapshot: typeof reference.native) => {
      await expect.poll(() => wordPlacements(editor)).toEqual(positionsFor(snapshot));
      await expect(page.locator('.section-page')).toHaveCount(
        Math.max(
          ...snapshot.paragraphs.flatMap((p) => [p.caret.page, ...p.characters.map((c) => c.page)]),
        ),
      );
    };
    await check(reference.steps.initial);
    await clickWordText(
      page,
      editor,
      reference.steps.initial.text.startsWith('before') ? 'before' : 'after',
    );
    await page.keyboard.press('Control+Home');
    const at = reference.steps.initial.text.search(/[\f\u000e]/);
    for (let i = 0; i < at; i++) await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Delete');
    if (reference.context === 'start') await page.keyboard.press('Control+Home');
    if (reference.context === 'selected')
      for (let i = 0; i < 3; i++) await page.keyboard.press('Shift+ArrowRight');
    if (name === 'column-single') {
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Column break', exact: true }).click();
      await page.getByRole('button', { name: 'Home', exact: true }).click();
    } else
      await page.keyboard.press(
        reference.steps.initial.text[at] === '\f' ? 'Control+Enter' : 'Control+Shift+Enter',
      );
    if (printRoute === 'keyboard') await page.keyboard.press('Control+p');
    if (printRoute === 'menu') {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      await page.getByRole('button', { name: 'Print / Save as PDF', exact: true }).click();
    }
    if (printRoute) {
      await expect
        .poll(() =>
          page.evaluate(() => (window as Window & { printSnapshot?: unknown }).printSnapshot),
        )
        .toEqual({
          text: reference.native.text.replace(/[\r\f\u000e]/g, ''),
          pages: Math.max(
            ...reference.native.paragraphs.flatMap((p) => [
              p.caret.page,
              ...p.characters.map((c) => c.page),
            ]),
          ),
        });
      await expect(page.locator('.word-fragment-print')).toHaveCount(0);
    }
    await check(reference.native);
    const selectionOffset = () =>
      editor.evaluate((root) => {
        const s = window.getSelection()!,
          range = document.createRange();
        range.setStart(root, 0);
        range.setEnd(s.anchorNode!, s.anchorOffset);
        const fragment = range.cloneContents();
        let n = 0;
        for (const p of fragment.querySelectorAll('p')) {
          n +=
            (p.textContent ?? '').length +
            p.querySelectorAll('[data-word-page-break],[data-word-column-break]').length;
          if (p !== fragment.lastChild) n++;
        }
        return n;
      });
    await expect.poll(selectionOffset).toBe(reference.native.selection.from);
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(editor.locator('p')).toHaveCount(1);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await check(reference.native);
    const reload = async () => {
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      await page.reload();
      await page.getByRole('button', { name: 'Recent files', exact: true }).click();
      await page
        .getByRole('button', { name: `Command ${name}`, exact: true })
        .first()
        .click();
    };
    const download = async (label: string) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page
        .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
        .click();
      const bytes = await fs.readFile((await (await pending).path())!);
      await fs.writeFile(`${root}/${label}.docx`, bytes);
      return bytes;
    };
    await reload();
    await check(reference.native);
    const positions = await wordPlacements(editor),
      lines = await wordLineOrigins(editor);
    const exported = await download('browser');
    const print = await page.pdf({ path: `${root}/browser-print.pdf`, preferCSSPageSize: true });
    const zoomPrints: { zoom: number; sourceHash?: string; editedHash?: string }[] = [];
    const printZooms = async (stage: 'source' | 'edited') => {
      if (!zoomPrint) return;
      await editor.focus();
      await page.keyboard.press('Control+End');
      const selected =
        (stage === 'source' ? reference.native.text : reference.steps.typed.text).length - 1;
      await expect.poll(selectionOffset).toBe(selected);
      for (const zoom of [50, 150, 100]) {
        await page.getByRole('button', { name: 'View', exact: true }).click();
        await page
          .getByRole('spinbutton', { name: 'Document zoom', exact: true })
          .fill(String(zoom));
        await page.getByRole('button', { name: 'Home', exact: true }).click();
        const calls = () =>
          page.evaluate(
            () => (window as Window & { printInvocations?: number }).printInvocations || 0,
          );
        const before = await calls();
        await page.keyboard.press('Control+p');
        try {
          await expect.poll(calls).toBe(before + 1);
        } catch (error) {
          await fs.writeFile(
            `${root}/print-timeout-${stage}-${zoom}.json`,
            JSON.stringify(
              await editor.evaluate(async (root) => {
                const samples = [];
                for (let i = 0; i < 12; i++) {
                  await new Promise(requestAnimationFrame);
                  samples.push({
                    html: root.innerHTML,
                    width: root.getBoundingClientRect().width,
                    paints: [
                      ...root.querySelectorAll<HTMLElement>(
                        '[data-word-baseline-paint],[data-word-run-leading]',
                      ),
                    ].map((span) => ({
                      text: span.textContent,
                      top: span.getBoundingClientRect().top,
                      parent: span.closest('p')!.getBoundingClientRect().top,
                      style: span.getAttribute('style'),
                    })),
                  });
                }
                return samples;
              }),
              null,
              2,
            ),
          );
          throw error;
        }
        await expect.poll(selectionOffset).toBe(selected);
        await expect(page.locator('.word-fragment-print')).toHaveCount(0);
        if (zoom === 100) continue;
        const bytes = await page.pdf({
          path: `${root}/browser-${stage === 'edited' ? 'edited-' : ''}print-zoom${zoom}.pdf`,
          preferCSSPageSize: true,
        });
        let row = zoomPrints.find((row) => row.zoom === zoom);
        if (!row) zoomPrints.push((row = { zoom }));
        row[stage === 'source' ? 'sourceHash' : 'editedHash'] = hash(bytes);
      }
    };
    await printZooms('source');
    await editor.focus();
    await page.keyboard.press('Control+End');
    await page.keyboard.insertText('!');
    await check(reference.steps.typed);
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await check(reference.native);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await check(reference.steps.typed);
    await reload();
    await check(reference.steps.typed);
    const editedPositions = await wordPlacements(editor),
      editedLines = await wordLineOrigins(editor);
    const edited = await download('browser-edited');
    const editedPrint = await page.pdf({
      path: `${root}/browser-edited-print.pdf`,
      preferCSSPageSize: true,
    });
    await printZooms('edited');
    await upload(`${name} exported`, exported);
    await check(reference.native);
    await upload(`${name} edited`, edited);
    await check(reference.steps.typed);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify(
        {
          passed: true,
          sourceHash: reference.sourceSha256,
          inputHash: hash(source),
          exportHash: hash(exported),
          editedHash: hash(edited),
          printHash: hash(print),
          editedPrintHash: hash(editedPrint),
          zoomPrints,
          positions,
          lines,
          editedPositions,
          editedLines,
          browser: page.context().browser()!.version(),
        },
        null,
        2,
      ) + '\n',
    );
  });
