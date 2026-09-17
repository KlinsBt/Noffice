import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import {
  clickWordText,
  wordPlacements,
  wordLineOrigins,
  waitWordLayout,
} from './word-pagination-helpers';
import { readFileSync } from 'node:fs';

// Every case owns its browser storage and artifact directory. Use the existing
// four-worker limit instead of leaving the remaining workers idle for this matrix.
test.describe.configure({ mode: 'parallel' });

const sectionNames = [
  'continuous',
  'continuous-dimensions',
  'columns-balanced',
  'columns-overflow',
  'continuous-columns',
  'continuous-from-columns',
  'keep-lines',
  'widow',
  ...Object.keys(
    JSON.parse(readFileSync('tests/fixtures/native-word-section-flow.json', 'utf8')).cases,
  ).filter((name) => /^(next-column|oddPage-|evenPage-)/.test(name)),
] as const;
const combinations = JSON.parse(
  readFileSync('tests/fixtures/native-word-forced-keep.json', 'utf8'),
) as typeof import('./fixtures/native-word-forced-keep.json');
const unequal = JSON.parse(
  readFileSync('tests/fixtures/native-word-unequal-flow.json', 'utf8'),
) as typeof import('./fixtures/native-word-unequal-flow.json');
for (const { name, family, prefix, transition } of [
  ...Object.entries(unequal.cases).map(([name]) => ({
    name,
    family: 'unequal',
    prefix: 'word-unequal-flow',
    transition: false,
  })),
  ...sectionNames.map((name) => ({
    name,
    family: 'section',
    prefix: 'word-section-flow',
    transition: false,
  })),
  ...sectionNames
    .filter((name) => /^(oddPage|evenPage)-lines-5(?:-large)?$/.test(name))
    .map((name) => ({ name, family: 'section', prefix: 'word-section-flow', transition: true })),
  ...Object.keys(combinations.cases).map((name) => ({
    name,
    family: 'forced',
    prefix: 'word-forced-keep',
    transition: false,
  })),
])
  test(`Word section flow ${transition ? 'transition-' : ''}${name}: editing, history, reload, print and DOCX`, async ({
    page,
  }) => {
    const root = `.local/word-${family}-acceptance/${transition ? 'transition-' : ''}${name}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(`tests/fixtures/${prefix}-${name}.docx`);
    const native = JSON.parse(
      await fs.readFile(`tests/fixtures/native-${prefix}.json`, 'utf8'),
    ) as typeof import('./fixtures/native-word-section-flow.json');
    const reference = native.cases[name as keyof typeof native.cases];
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    const pdfDownloads =
      family === 'unequal' || (family === 'forced' && name === 'keep-column-mid');
    if (pdfDownloads) await page.context().grantPermissions(['local-fonts']);
    const downloadPdf = async (stage: string) => {
      if (!pdfDownloads) return undefined;
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: 'PDF file', exact: true }).click();
      const bytes = await fs.readFile((await (await pending).path())!);
      expect(bytes.subarray(0, 5).toString()).toBe('%PDF-');
      await fs.writeFile(`${root}/download-${stage}.pdf`, bytes);
      return hash(bytes);
    };
    expect(hash(source)).toBe(reference.sourceSha256);
    const expected = reference.native.paragraphs.map((p) =>
      p.characters.map((c) => ({ text: c.text.replaceAll('\v', '\n'), page: c.page })),
    );
    const wrapped = family === 'unequal' && name.endsWith('-wrapped');
    const expectedLines = reference.native.paragraphs.map((p) =>
      p.characters
        .filter(
          (c, i) =>
            !['\f', '\u000e'].includes(c.text) &&
            (i === 0 ||
              ['\v', '\f', '\u000e'].includes(p.characters[i - 1].text) ||
              (wrapped &&
                (c.page !== p.characters[i - 1].page ||
                  Math.abs(c.y - p.characters[i - 1].y) > 0.01 ||
                  c.x < p.characters[i - 1].x - 0.01))),
        )
        .map((c) => ({ page: c.page, x: c.x })),
    );
    if (family === 'unequal') {
      const printed = unequal.cases[name as keyof typeof unequal.cases].native.pdfLines;
      expect(expectedLines.flat()).toHaveLength(printed.length);
      let index = 0;
      for (const paragraph of expectedLines)
        for (const line of paragraph) {
          expect(line.page).toBe(printed[index].page);
          // COM screen coordinates have a different quantization from the PDF
          // used by the print gate. Compare physical origins to that same oracle.
          line.x = printed[index++].x;
        }
    }
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
    await upload(`Flow ${name}`, source);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    const check = async () => {
      await waitWordLayout(editor);
      await expect.poll(() => wordPlacements(editor)).toEqual(expected);
      if ('pdfPages' in reference.native)
        await expect
          .poll(() =>
            page.locator('.section-page').evaluateAll((pages) =>
              pages.map((page) => ({
                width: parseFloat(getComputedStyle(page).width) * 0.75,
                height: parseFloat(getComputedStyle(page).height) * 0.75,
              })),
            ),
          )
          .toEqual(reference.native.pdfPages.map(({ width, height }) => ({ width, height })));
      if (family === 'unequal') {
        await expect
          .poll(async () => {
            const actual = await wordLineOrigins(editor);
            return (
              actual.length === expectedLines.length &&
              actual.every(
                (lines, p) =>
                  lines.length === expectedLines[p].length &&
                  lines.every(
                    (line, i) =>
                      line.page === expectedLines[p][i].page &&
                      Math.abs(line.x - expectedLines[p][i].x) <= 0.15,
                  ),
              )
            );
          })
          .toBe(true);
        const rows = unequal.cases[name as keyof typeof unequal.cases].native.pdfPages;
        const rules = rows.flatMap((p, page) => p.paths.map((rule) => ({ page, ...rule })));
        await expect(page.locator('.section-pages .word-column-separator')).toHaveCount(
          rules.length,
        );
        const actualRules = await page
          .locator('.section-pages .word-column-separator')
          .evaluateAll((els) =>
            els.map((el) => {
              const surface = el.parentElement!,
                parent = surface.getBoundingClientRect(),
                rect = el.getBoundingClientRect();
              const scale = parent.width / parseFloat(getComputedStyle(surface).width);
              return {
                page: [...surface.parentElement!.children].indexOf(surface),
                x: (((rect.left + rect.right) / 2 - parent.left) * 0.75) / scale,
                top: ((rect.top - parent.top) * 0.75) / scale,
                bottom: ((rect.bottom - parent.top) * 0.75) / scale,
              };
            }),
          );
        for (let i = 0; i < rules.length; i++) {
          expect(actualRules[i].page).toBe(rules[i].page);
          expect(
            Math.abs(actualRules[i].x - (rules[i].left + rules[i].right) / 2),
          ).toBeLessThanOrEqual(0.15);
          expect(Math.abs(actualRules[i].top - rules[i].top)).toBeLessThanOrEqual(0.15);
          expect(Math.abs(actualRules[i].bottom - rules[i].bottom)).toBeLessThanOrEqual(0.15);
        }
      } else
        await expect
          .poll(async () =>
            (await wordLineOrigins(editor)).map((p) => p.map((l) => ({ page: l.page, x: l.x }))),
          )
          .toEqual(expectedLines);
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
      lines = await wordLineOrigins(editor);
    for (let i = 0; i < 5; i++)
      await page.getByRole('button', { name: 'Zoom out', exact: true }).click();
    await check();
    const firstLabels = reference.native.paragraphs[0].characters
      .map((c) => c.text)
      .join('')
      .split(wrapped ? /[ \v]+/ : '\v');
    const target = wrapped
      ? unequal.cases[name as keyof typeof unequal.cases].native.pdfLines[4].text
          .trim()
          .split(' ')[0]
      : firstLabels[Math.min(4, firstLabels.length - 1)];
    await clickWordText(page, editor, target);
    await page.keyboard.press(wrapped ? 'Home' : 'End');
    await page.keyboard.insertText('!');
    await expect(editor.locator('p').first()).toContainText(wrapped ? `!${target}` : `${target}!`);
    await page.keyboard.press('Control+z');
    await check();
    for (let i = 0; i < 5; i++)
      await page.getByRole('button', { name: 'Zoom in', exact: true }).click();
    await check();
    const last = firstLabels.at(-1)!;
    await clickWordText(page, editor, last);
    await page.keyboard.press('End');
    const added = transition
      ? Array.from({ length: 6 }, (_, i) => `s1l${String(i + 6).padStart(2, '0')}`)
      : ['added'];
    for (const text of added) {
      await page.keyboard.press('Shift+Enter');
      await page.keyboard.insertText(text);
    }
    await expect(editor.locator('p').first()).toContainText(`${last}${added.join('')}`);
    const checkTransition = async () => {
      if (!transition) return;
      const target = native.cases[name.replace('lines-5', 'lines-11') as keyof typeof native.cases];
      await expect
        .poll(() => wordPlacements(editor))
        .toEqual(
          target.native.paragraphs.map((p) =>
            p.characters.map((c) => ({ text: c.text.replaceAll('\v', '\n'), page: c.page })),
          ),
        );
      if ('pdfPages' in target.native)
        await expect(page.locator('.section-page')).toHaveCount(target.native.pdfPages.length);
    };
    await checkTransition();
    await waitWordLayout(editor);
    const editedPositions = await wordPlacements(editor),
      editedLines = await wordLineOrigins(editor);
    // Controls need a real caret/page position as well as painted text.
    // display:contents on a reflow BR once passed history checks with page 0.
    expect(editedPositions.flat().every((character) => character.page > 0)).toBe(true);
    // A single typing history event contains the entered line break and text.
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await check();
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
    await checkTransition();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page
      .getByRole('button', { name: `Flow ${name}`, exact: true })
      .first()
      .click();
    await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
    await checkTransition();
    const edited = await download('browser-edited');
    const editedPDF = await page.pdf({
      path: `${root}/browser-edited-print.pdf`,
      preferCSSPageSize: true,
    });
    await expect(page.locator('.word-fragment-print')).toHaveCount(0);
    const editedPdfDownloadHash = await downloadPdf('edited');
    await upload(`Flow ${name} edited`, edited);
    await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
    await checkTransition();
    await clickWordText(page, editor, added.at(-1)!);
    await page.keyboard.press('End');
    for (let i = 0; i < added.reduce((n, text) => n + text.length + 1, 0); i++)
      await page.keyboard.press('Backspace');
    await check();
    const restored = await download('browser');
    const pdf = await page.pdf({ path: `${root}/browser-print.pdf`, preferCSSPageSize: true });
    const pdfDownloadHash = await downloadPdf('source');
    await upload(`Flow ${name} restored`, restored);
    await check();
    const stored = await page.evaluate(async () => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const r = indexedDB.open('noffice-workspace');
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
      try {
        return JSON.stringify(
          await new Promise((resolve, reject) => {
            const r = db.transaction('files').objectStore('files').getAll();
            r.onsuccess = () => resolve(r.result);
            r.onerror = () => reject(r.error);
          }),
        );
      } finally {
        db.close();
      }
    });
    expect(stored).not.toMatch(
      /word-fragment-print|fragment-left|fragment-shift|data-fragment-column|data-word-reflow-top/,
    );
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
          pdfDownloadHash,
          editedPdfDownloadHash,
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
