import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { wordStoryBuildHash } from './word-story-artifacts';

test('mixed-size soft lines preserve native wrapping and advances through spacing edits and exports', async ({
  page,
}) => {
  test.setTimeout(90000);
  const run = process.env.NOFFICE_MIXED_WRAP_RUN;
  if (run && !/^browser-v\d+$/.test(run)) throw Error('Invalid mixed-wrap evidence folder');
  const root = '.local/word-mixed-wrap' + (run ? '/' + run : '');
  const glyphRoot = '.local/word-glyph' + (run ? '/' + run : '');
  const prints: Record<string, string> = {};
  await fs.mkdir(glyphRoot, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-mixed-wrap.docx');
  const native = JSON.parse(
    await fs.readFile('tests/fixtures/native-word-mixed-wrap.json', 'utf8'),
  );
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  expect(hash(source)).toBe(native.sourceSha256);
  await fs.mkdir(root, { recursive: true });
  await fs.writeFile(`${root}/source.docx`, source);
  await page.setViewportSize({ width: 1500, height: 1200 });
  await page.goto('/');
  const upload = async (name: string, buffer: Buffer) => {
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `${name}.docx`,
      buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  await upload('Mixed wrap', source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const paragraphs = editor.locator('p');
  await expect(paragraphs).toHaveCount(4);
  const measure = async () =>
    paragraphs.evaluateAll((ps) =>
      ps.slice(1, 3).map((p) => {
        const bounds = p.getBoundingClientRect();
        const scale = bounds.width / parseFloat(getComputedStyle(p).width);
        const lines: { words: string[]; baseline: number }[] = [];
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        const pieces: { text: Text; start: number; ascent: number }[] = [];
        let content = '';
        while (walker.nextNode()) {
          const text = walker.currentNode as Text;
          if (!text.data) continue;
          const style = getComputedStyle(text.parentElement!);
          const probe = document.createElement('span');
          probe.style.cssText =
            'all:initial;position:fixed;left:-100000px;top:0;visibility:hidden;display:inline-block;white-space:pre;line-height:normal';
          probe.style.fontFamily = style.fontFamily;
          probe.style.zoom = String(scale);
          probe.style.fontSize = style.fontSize;
          probe.style.fontWeight = style.fontWeight;
          probe.style.fontStyle = style.fontStyle;
          const probeText = document.createTextNode('Hg');
          const marker = document.createElement('span');
          marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline';
          probe.append(probeText, marker);
          document.body.append(probe);
          const range = document.createRange();
          range.selectNodeContents(probeText);
          const rangeAscent =
            marker.getBoundingClientRect().top - range.getBoundingClientRect().top;
          probe.remove();
          pieces.push({ text, start: content.length, ascent: rangeAscent });
          content += text.data;
        }
        // Edit provenance and paint spans can split a word into several text
        // nodes. Match semantic words across those boundaries and measure every
        // piece, so neither missing text nor a displaced fragment is hidden.
        for (const match of content.matchAll(/(?:alpha|beta)\d+/g)) {
          const start = match.index!, end = start + match[0].length;
          const baselines = pieces.filter(piece => piece.start < end && piece.start + piece.text.length > start)
            .map(piece => {
              const range = document.createRange();
              range.setStart(piece.text, Math.max(0, start - piece.start));
              range.setEnd(piece.text, Math.min(piece.text.length, end - piece.start));
              return ((range.getBoundingClientRect().top + piece.ascent - bounds.top) * 0.75) / scale;
            });
          const baseline = baselines[0];
          if (!Number.isFinite(baseline) || baselines.some(value => Math.abs(value - baseline) >= 0.15))
            throw Error('A semantic word has missing or displaced paint fragments');
          if (!lines.length || Math.abs(lines.at(-1)!.baseline - baseline) > 0.5)
            lines.push({ words: [], baseline });
          lines.at(-1)!.words.push(match[0]);
        }
        const boxes = [...p.querySelectorAll('[data-word-softline-strut]')].map((strut) => {
          const box = strut.getBoundingClientRect();
          return {
            top: ((box.top - bounds.top) * 0.75) / scale,
            height: (box.height * 0.75) / scale,
          };
        });
        return { height: (bounds.height * 0.75) / scale, lines, boxes, html: p.outerHTML };
      }),
    );
  const check = async (stage: string) => {
    await expect(paragraphs.nth(1)).toHaveAttribute('data-word-wrapped-leading', 'true');
    await expect
      .poll(async () => (await measure()).map((p) => p.lines.map((l) => l.words)))
      .toEqual(native.stages[stage].map((p: any) => p.lines.map((l: any) => l.words)));
    await fs.writeFile(
      `${root}/browser-${stage}-pending.json`,
      JSON.stringify(await measure(), null, 2),
    );
    for (let i = 0; i < 2; i++)
      await expect
        .poll(async () => Math.abs((await measure())[i].height - native.stages[stage][i].height))
        .toBeLessThan(0.15);
    await expect
      .poll(async () => {
        const measured = await measure();
        return Math.max(
          ...measured.flatMap((p, i) => {
            const differences = p.lines.map(
              (line, j) => line.baseline - native.stages[stage][i].lines[j].baseline,
            );
            return differences.flatMap((difference) => [
              Math.abs(difference),
              Math.abs(difference - differences[0]),
            ]);
          }),
        );
      })
      .toBeLessThan(0.15);
    const result = await measure();
    await fs.writeFile(`${root}/browser-${stage}-metrics.json`, JSON.stringify(result, null, 2));
    for (let i = 0; i < 2; i++)
      expect(Math.abs(result[i].height - native.stages[stage][i].height)).toBeLessThan(0.15);
    for (let i = 0; i < 2; i++) {
      expect(result[i].boxes).toHaveLength(native.stages[stage][i].lines.length);
      for (let j = 0; j < result[i].boxes.length; j++) {
        expect(
          Math.abs(result[i].boxes[j].top - native.stages[stage][i].lines[j].boxTop),
        ).toBeLessThan(0.15);
        expect(
          Math.abs(result[i].boxes[j].height - native.stages[stage][i].lines[j].boxHeight),
        ).toBeLessThan(0.15);
      }
    }
    return result;
  };
  const capture = async (stage: string) => {
    // Source identities survive print cloning; arbitrary DOM attributes may be
    // discarded when ProseMirror redraws the editable paragraph.
    const ids = await paragraphs.evaluateAll(ps => ps.slice(1, 3).map(p => p.getAttribute('data-source-paragraph')));
    expect(ids.every(id => id !== null)).toBe(true);
    const style = await page.addStyleTag({
      content:
        `[data-source-paragraph=${JSON.stringify(ids[0])}]{background:#00ffff!important;print-color-adjust:exact}` +
        `[data-source-paragraph=${JSON.stringify(ids[1])}]{background:#ff00ff!important;print-color-adjust:exact}`,
    });
    try {
      await check(stage);
      const bytes = await page.pdf({
        path: `${glyphRoot}/${stage}.pdf`,
        printBackground: true,
        preferCSSPageSize: true,
      });
      prints[stage] = hash(bytes);
    } finally {
      await style.evaluate((el) => el.parentNode?.removeChild(el));
      await check(stage);
    }
  };
  const initial = await check('source');
  await capture('source');
  await editor.screenshot({ path: `${root}/browser-source.png` });
  const select = async (word: string, offset = 0, length = word.length) => {
    await paragraphs.nth(1).evaluate(
      (p, args) => {
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        while (walker.nextNode()) {
          const text = walker.currentNode as Text;
          const index = text.data.indexOf(args.word);
          if (index < 0) continue;
          const range = document.createRange();
          range.setStart(text, index + args.offset);
          range.setEnd(text, index + args.offset + args.length);
          const selection = getSelection()!;
          selection.removeAllRanges();
          selection.addRange(range);
          (p.closest('[contenteditable]') as HTMLElement).focus();
          document.dispatchEvent(new Event('selectionchange'));
          return;
        }
        throw new Error('Missing selection word');
      },
      { word, offset, length },
    );
  };
  const stages: Record<string, unknown> = {};
  const exports: Record<string, string> = {};
  for (const stage of ['double', 'minimum', 'single', 'size', 'typing']) {
    await upload(`Mixed ${stage} edit`, source);
    if (stage === 'size' || stage === 'typing') {
      await select(
        stage === 'size' ? 'alpha5' : 'alpha1',
        stage === 'typing' ? 5 : 0,
        stage === 'typing' ? 0 : 6,
      );
      if (stage === 'size') {
        await expect.poll(() => page.evaluate(() => getSelection()?.toString())).toBe('alpha5');
        const size = page.getByRole('spinbutton', { name: 'Font size', exact: true });
        await size.fill('20');
        await size.press('Enter');
      } else await page.keyboard.type('0');
      await check(stage);
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await check('source');
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
    } else {
      for (let index = 1; index <= 2; index++) {
        await paragraphs.nth(index).click();
        const spacing = page.getByRole('combobox', { name: 'Line spacing', exact: true });
        if (stage === 'minimum') {
          await spacing.selectOption('custom');
          const dialog = page.getByRole('dialog', { name: 'Line spacing', exact: true });
          await dialog
            .getByRole('combobox', { name: 'Spacing rule', exact: true })
            .selectOption('atLeast');
          const amount = dialog.getByRole('spinbutton', { name: 'Spacing amount', exact: true });
          await amount.fill('0');
          await expect(dialog.getByRole('button', { name: 'Apply', exact: true })).toBeDisabled();
          await amount.fill('50');
          await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
        } else await spacing.selectOption(stage === 'double' ? '2' : '1');
        await page.getByRole('button', { name: 'Undo', exact: true }).click();
        await expect
          .poll(async () =>
            Math.abs((await measure())[index - 1].height - initial[index - 1].height),
          )
          .toBeLessThan(0.05);
        await page.getByRole('button', { name: 'Redo', exact: true }).click();
      }
    }
    stages[stage] = await check(stage);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: `Mixed ${stage} edit`, exact: true }).click();
    await check(stage);
    await capture(stage);
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    exports[stage] = hash(bytes);
    await fs.writeFile(`${root}/browser-${stage}.docx`, bytes);
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    expect(xml).not.toContain('strut');
    expect(xml).not.toContain('measured');
    await upload(`Mixed ${stage} reimport`, bytes);
    await check(stage);
  }
  await upload('Mixed wrap recovery', source);
  await select('alpha5');
  await page.getByRole('button', { name: 'Bold', exact: true }).click();
  await expect(paragraphs.nth(1)).not.toHaveAttribute('data-word-wrapped-leading');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await check('source');
  await select('alpha5');
  const unmeasuredSize = page.getByRole('spinbutton', { name: 'Font size', exact: true });
  await unmeasuredSize.fill('21');
  await unmeasuredSize.press('Enter');
  await expect(paragraphs.nth(1).locator('[style*="--word-baseline-shift"]')).toHaveCount(0);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await check('source');
  await select('alpha1', 0, 0);
  await page.keyboard.insertText('A'.repeat(5000));
  await expect(paragraphs.nth(1)).not.toHaveAttribute('data-word-wrapped-leading');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await check('source');
  // Width changes must remeasure the original soft lines; restoring width must
  // restore the native fixture's identities without a document/history edit.
  await editor.evaluate((el) => (el.style.width = '200px'));
  await expect.poll(async () => (await measure())[0].lines.length).toBeGreaterThan(6);
  await editor.evaluate((el) => el.style.removeProperty('width'));
  await check('source');
  for (let i = 0; i < 5; i++)
    await page.getByRole('button', { name: 'Zoom out', exact: true }).click();
  const zoomed = await check('source');
  const hit = await paragraphs.nth(1).evaluate((p) => {
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const text = walker.currentNode as Text;
      const index = text.data.indexOf('alpha5');
      if (index < 0) continue;
      const range = document.createRange();
      range.setStart(text, index);
      range.setEnd(text, index + 1);
      const bounds = range.getBoundingClientRect();
      return { x: bounds.left + bounds.width / 2, y: bounds.top + bounds.height / 2 };
    }
    throw new Error('Missing second soft line');
  });
  await page.mouse.click(hit.x, hit.y);
  await page.keyboard.press('Home');
  await page.keyboard.insertText('X');
  await expect(paragraphs.nth(1)).toContainText('Xalpha5');
  await page.keyboard.press('Control+z');
  await expect(paragraphs.nth(1)).not.toContainText('Xalpha5');
  await check('source');
  for (let i = 0; i < 5; i++)
    await page.getByRole('button', { name: 'Zoom in', exact: true }).click();
  await check('source');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const stored = await page.evaluate(async () => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('noffice-workspace');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const records = await new Promise<unknown[]>((resolve, reject) => {
      const request = db.transaction('files').objectStore('files').getAll();
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return JSON.stringify(records);
  });
  expect(stored).not.toContain('softline-strut');
  expect(stored).not.toContain('word-wrapped-leading');
  expect(stored).not.toContain('word-baseline-shift');
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      { passed: true, sourceSha256: hash(source), exports, prints, initial, stages, zoomed,
        buildHash: await wordStoryBuildHash(),
        testHash: hash(await fs.readFile('tests/word-mixed-wrap.spec.ts')) },
      null,
      2,
    ),
  );
});
