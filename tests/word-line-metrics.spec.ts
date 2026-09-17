import { test, expect } from '@playwright/test';
import { lineMetricsFixture, fontCascadeFixture } from '../scripts/word-line-metrics-fixture.mjs';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import type { OfficeFile } from '../src/model';

for (const cascade of [false, true])
  test(`Word ${cascade ? 'font cascade' : 'fallback fonts'} and spacing survive dialog history and actual exports`, async ({
    page,
  }) => {
    await page.setViewportSize({ width: 1500, height: 1200 });
    const root = cascade ? '.local/word-font-cascade' : '.local/word-line-metrics';
    await fs.mkdir(root, { recursive: true });
    const source = await (cascade ? fontCascadeFixture() : lineMetricsFixture());
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    await fs.writeFile(`${root}/source.docx`, source);
    await fs.writeFile(`${root}/source.json`, JSON.stringify({ sha256: hash(source) }));
    await page.goto('/');
    const upload = async (name: string, buffer: Buffer) => {
      await page.locator('input[type=file][multiple]').setInputFiles({
        name: `${name}.docx`,
        buffer,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
    };
    await upload('Line metrics', source);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    const target = editor.locator('p').nth(1);
    await expect(editor.locator('p')).toHaveCount(6);
    await expect(editor.locator('p').first()).toHaveCSS('width', '280px');
    await expect(target).toHaveAttribute('data-word-line-rule', 'atLeast');
    // Word omits an explicit auto=240 attribute when saving its default single spacing.
    expect(
      await editor
        .locator('p')
        .nth(2)
        .evaluate((p) => p.style.lineHeight),
    ).toBe('1');
    // The wrapped layout uses per-line struts, so the paragraph's computed
    // line-height is no longer its physical advance. Verify native geometry.
    let wrappedAdvances: number[] = [];
    await expect
      .poll(async () => {
        wrappedAdvances = await editor
          .locator('p')
          .nth(4)
          .evaluate((paragraph) => {
            const scale =
              paragraph.getBoundingClientRect().width /
              parseFloat(getComputedStyle(paragraph).width);
            const tops: number[] = [];
            const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
            while (walker.nextNode()) {
              const text = walker.currentNode as Text;
              if (text.parentElement?.closest('.ProseMirror-widget')) continue;
              for (let index = 0; index < text.length; index++) {
                if (/\s/.test(text.data[index])) continue;
                const range = document.createRange();
                range.setStart(text, index);
                range.setEnd(text, index + 1);
                const top = range.getBoundingClientRect().top / scale;
                if (!tops.some((value) => Math.abs(value - top) < 1)) tops.push(top);
              }
            }
            tops.sort((a, b) => a - b);
            return tops.slice(1).map((top, index) => top - tops[index]);
          });
        return (
          wrappedAdvances.length > 0 &&
          wrappedAdvances.every((advance) => Math.abs(advance - 23) <= 0.2)
        );
      })
      .toBe(true);
    if (cascade) await expect(editor.locator('p').nth(3)).toHaveCSS('height', '46px');
    const text = await editor.textContent();
    const measurements = await editor.locator('p').evaluateAll((ps) =>
      ps.map((p) => ({
        height: p.getBoundingClientRect().height,
        lineHeight: getComputedStyle(p).lineHeight,
        rule: p.getAttribute('data-word-line-rule'),
        // Wrapped paragraphs suppress their own CSS strut with font-size:0.
        // The paragraph mark remains authored in points; verify painted runs below.
        fontSize: parseFloat(p.style.fontSize),
        fontFamily: getComputedStyle(p).fontFamily,
        runs: [...p.querySelectorAll('span')]
          .filter((s) => s.textContent && !s.querySelector('span'))
          .map((s) => ({
            text: s.textContent,
            size: parseFloat(getComputedStyle(s).fontSize) * 0.75,
            family: getComputedStyle(s).fontFamily,
          })),
      })),
    );
    for (const [index, p] of measurements.entries()) {
      expect(p.fontSize).toBeCloseTo(
        cascade && index === 1 ? 24 : cascade && index === 3 ? 30 : 10,
        2,
      );
      expect(p.fontFamily).toBe('Arial');
      for (const run of p.runs)
        expect(run.size).toBeCloseTo(
          run.text === 'Large line'
            ? 20
            : cascade && index === 1
              ? run.text === 'Small again'
                ? 18
                : 14
              : 10,
          2,
        );
    }
    expect(measurements[1].height).toBeGreaterThan(measurements[0].height);
    // Independent Word source snapshots: 30pt empty mark/minimum 18pt advances
    // 34.5pt; homogeneous 10pt automatic 1.5 advances 17.25pt per line.
    await target.click();
    await page.keyboard.press('Home');
    const menu = page.getByRole('combobox', { name: 'Line spacing', exact: true });
    await expect(menu.locator('option:checked')).toHaveText('At least 12pt');
    await menu.selectOption('custom');
    const dialog = page.getByRole('dialog', { name: 'Line spacing', exact: true });
    const amount = dialog.getByRole('spinbutton', { name: 'Spacing amount', exact: true });
    await expect(amount).toHaveValue('12');
    await amount.fill('0');
    await expect(dialog.getByRole('button', { name: 'Apply', exact: true })).toBeDisabled();
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
    await expect(editor).toBeFocused();
    await expect(target).toHaveCSS('line-height', '16px');
    const exports: Record<string, string> = {};
    for (const [name, rule, value, css] of [
      ['minimum', 'atLeast', '18', '24px'],
      ['exact', 'exact', '18', '24px'],
      ['auto', 'auto', '1.5', cascade ? '48px' : '20px'],
    ]) {
      await target.click();
      await page.keyboard.press('Home');
      const previous = await target.getAttribute('style');
      const previousRule = await target.getAttribute('data-word-line-rule');
      await menu.selectOption('custom');
      await dialog.getByRole('combobox', { name: 'Spacing rule', exact: true }).selectOption(rule);
      await amount.fill(value);
      await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      await expect(dialog).toHaveCount(0);
      await expect(target).toHaveCSS('line-height', css);
      if (rule === 'atLeast')
        await expect(target).toHaveAttribute('data-word-line-rule', 'atLeast');
      else await expect(target).not.toHaveAttribute('data-word-line-rule');
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await expect(target).toHaveAttribute('style', previous!);
      expect(await target.getAttribute('data-word-line-rule')).toBe(previousRule);
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await expect(target).toHaveCSS('line-height', css);
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      const filename = await page
        .getByRole('textbox', { name: 'File name', exact: true })
        .inputValue();
      await page.reload();
      await page.getByRole('button', { name: 'Recent files', exact: true }).click();
      await page.getByRole('button', { name: filename, exact: true }).click();
      await expect(target).toHaveCSS('line-height', css);
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page
        .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
        .click();
      const output = await fs.readFile((await (await pending).path())!);
      await fs.writeFile(`${root}/browser-${name}.docx`, output);
      exports[name] = hash(output);
      const zip = await JSZip.loadAsync(output);
      const xml = await zip.file('word/document.xml')!.async('string');
      expect(xml).toContain(`w:lineRule="${rule}"`);
      await upload(`Line metrics ${name}`, output);
      await expect(target).toHaveCSS('line-height', css);
      expect(await editor.textContent()).toBe(text);
    }
    if (cascade) {
      const empty = editor.locator('p').nth(3);
      await empty.click();
      await expect(page.getByRole('spinbutton', { name: 'Font size', exact: true })).toHaveValue(
        '30',
      );
      await page.keyboard.insertText('Mark text');
      await expect(empty).toHaveText('Mark text');
      expect(
        await empty
          .locator('span')
          .first()
          .evaluate((el) => parseFloat(getComputedStyle(el).fontSize) * 0.75),
      ).toBeCloseTo(30, 2);
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await expect(empty).toHaveText('');
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await expect(empty).toHaveText('Mark text');
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      await page.reload();
      await page.getByRole('button', { name: 'Recent files', exact: true }).click();
      await page.getByRole('button', { name: 'Line metrics auto', exact: true }).click();
      await expect(empty).toHaveText('Mark text');
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page
        .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
        .click();
      const bytes = await fs.readFile((await (await pending).path())!);
      await fs.writeFile(`${root}/browser-typing.docx`, bytes);
      exports.typing = hash(bytes);
      await upload('Paragraph font typing', bytes);
      await expect(empty).toHaveText('Mark text');
      expect(
        await empty
          .locator('span')
          .first()
          .evaluate((el) => parseFloat(getComputedStyle(el).fontSize) * 0.75),
      ).toBeCloseTo(30, 2);
    }
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify(
        {
          passed: true,
          wrappedAdvances,
          sourceHash: hash(source),
          exports,
          measurements,
          browser: page.context().browser()!.version(),
        },
        null,
        2,
      ),
    );
  });

test('legacy minimum spacing hydrates before editing without overwriting saved text or storage revisions', async ({
  page,
}) => {
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Legacy spacing.docx',
    buffer: await lineMetricsFixture(),
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Legacy spacing',
  );
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const stored = (strip: boolean) =>
    page.evaluate(
      (strip) =>
        new Promise<{ revision: number; version?: number; fonts?: number }>((resolve, reject) => {
          const request = indexedDB.open('noffice-workspace');
          request.onerror = () => reject(request.error);
          request.onsuccess = () => {
            const db = request.result,
              tx = db.transaction('files', strip ? 'readwrite' : 'readonly'),
              store = tx.objectStore('files');
            const all = store.getAll();
            let result: { revision: number; version?: number; fonts?: number };
            all.onsuccess = () => {
              const file = all.result.find(
                (f: OfficeFile) => f.name === 'Legacy spacing',
              ) as OfficeFile;
              if (file.content.kind !== 'word') {
                tx.abort();
                return;
              }
              if (strip) {
                delete file.content.lineSpacingVersion;
                delete file.content.fontMetricsVersion;
                const doc = new DOMParser().parseFromString(file.content.html, 'text/html');
                const paragraphs = doc.querySelectorAll('p');
                for (const p of paragraphs) {
                  p.style.removeProperty('font-size');
                  p.style.removeProperty('font-family');
                }
                for (const span of doc.querySelectorAll('span'))
                  if (span.style.fontSize === '10pt') span.style.removeProperty('font-size');
                paragraphs[1].style.removeProperty('line-height');
                paragraphs[1].removeAttribute('data-word-line-rule');
                paragraphs[1].append(' saved edit');
                paragraphs[3].style.lineHeight = '1.5';
                paragraphs[3].removeAttribute('data-word-line-rule');
                file.content.html = doc.body.innerHTML;
                store.put(file);
              }
              result = {
                revision: file.revision,
                version: file.content.lineSpacingVersion,
                fonts: file.content.fontMetricsVersion,
              };
            };
            tx.oncomplete = () => {
              db.close();
              resolve(result);
            };
            tx.onerror = tx.onabort = () => {
              db.close();
              reject(tx.error);
            };
          };
        }),
      strip,
    );
  const before = await stored(true);
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Legacy spacing', exact: true }).click();
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(editor.locator('p').nth(1)).toContainText('saved edit');
  expect(
    await editor
      .locator('p')
      .nth(1)
      // The measured mixed-line container has a zero strut; the hydrated
      // paragraph-mark size remains semantic, and actual runs paint their sizes.
      .evaluate((p) => parseFloat(p.style.fontSize) * (p.style.fontSize.endsWith('pt') ? 1 : 0.75)),
  ).toBeCloseTo(10, 2);
  const glyphSize = (word: string) =>
    editor
      .locator('p')
      .nth(1)
      .evaluate((p, word) => {
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) {
          const text = walker.currentNode as Text,
            offset = text.data.indexOf(word);
          if (offset < 0) continue;
          const range = document.createRange();
          range.setStart(text, offset);
          range.setEnd(text, offset + word.length);
          const box = range.getBoundingClientRect();
          if (!(box.width > 0 && box.height > 0)) throw Error('Hydrated glyph is not visible.');
          return parseFloat(getComputedStyle(text.parentElement!).fontSize) * 0.75;
        }
        throw Error(`Missing hydrated word ${word}`);
      }, word);
  expect(await glyphSize('Small')).toBeCloseTo(10, 2);
  expect(await glyphSize('Large')).toBeCloseTo(20, 2);
  await expect(editor.locator('p').nth(1)).toHaveAttribute('data-word-line-rule', 'atLeast');
  expect(
    await editor
      .locator('p')
      .nth(1)
      .evaluate(
        (p) => parseFloat(p.style.lineHeight) * (p.style.lineHeight.endsWith('pt') ? 1 : 0.75),
      ),
  ).toBeCloseTo(12, 2);
  const advances = await editor
    .locator('p')
    .nth(1)
    .locator('[data-word-line-from]')
    .evaluateAll((lines) =>
      lines.map((line) => parseFloat((line as HTMLElement).style.height) * 0.75),
    );
  expect(advances.length).toBeGreaterThanOrEqual(3);
  expect(advances.every((height) => height >= 12)).toBe(true);
  await expect(editor.locator('p').nth(3)).not.toHaveAttribute('data-word-line-rule');
  expect(await stored(false)).toEqual(before);
  await editor.locator('p').nth(1).click();
  await page.keyboard.press('End');
  await page.keyboard.insertText('!');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor.locator('p').nth(1)).not.toContainText('!');
  await expect(editor.locator('p').nth(1)).toHaveAttribute('data-word-line-rule', 'atLeast');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const after = await stored(false);
  expect(after.revision).toBeGreaterThan(before.revision);
  expect(after.version).toBe(1);
  expect(after.fonts).toBe(1);
});
