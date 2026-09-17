import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

test('uniform wrapped and hard-break text keeps leading placement through history and exports', async ({
  page,
}) => {
  const root = '.local/word-wrapped-leading';
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-wrapped-leading.docx');
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  const oracle = JSON.parse(
    (await fs.readFile('tests/fixtures/native-word-wrapped-leading.json', 'utf8')).replace(
      /^\uFEFF/,
      '',
    ),
  );
  expect(hash(source)).toBe(oracle.sourceSha256);
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
  await upload('Wrapped leading', source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const paragraphs = editor.locator('p');
  await expect(paragraphs).toHaveCount(4);
  await expect(paragraphs.nth(1)).toHaveCSS('width', '280px');
  const measure = (index: number) =>
    paragraphs.nth(index).evaluate((p) => {
      const bounds = p.getBoundingClientRect();
      const scale = bounds.width / parseFloat(getComputedStyle(p).width);
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
      const lines: { from: number; text: string; top: number }[] = [];
      let offset = 0;
      while (walker.nextNode()) {
        const node = walker.currentNode;
        if (node.nodeName === 'BR') {
          offset++;
          continue;
        }
        if (node.nodeType !== Node.TEXT_NODE) continue;
        for (let i = 0; i < node.textContent!.length; i++) {
          const char = node.textContent![i];
          if (/\s/.test(char)) continue;
          const range = document.createRange();
          range.setStart(node, i);
          range.setEnd(node, i + 1);
          const top = ((range.getBoundingClientRect().top - bounds.top) * 0.75) / scale;
          if (!lines.length || Math.abs(lines.at(-1)!.top - top) > 0.5)
            lines.push({ from: offset + i, top, text: '' });
          lines.at(-1)!.text += char;
        }
        offset += node.textContent!.length;
      }
      return { height: (bounds.height * 0.75) / scale, scale, lines };
    });
  const nativeLines = [1, 2].map((index) => {
    const lines: { from: number; text: string; y: number }[] = [];
    oracle.source[index].chars.forEach((c: { text: string; y: number }, from: number) => {
      if (/\s/.test(c.text)) return;
      if (!lines.length || lines.at(-1)!.y !== c.y) lines.push({ from, text: '', y: c.y });
      lines.at(-1)!.text += c.text;
    });
    return lines.map(({ from, text }) => ({ from, text }));
  });
  await expect.poll(async () => (await measure(1)).height).toBeCloseTo(7 * 17.25, 1);
  const initial = await Promise.all([measure(1), measure(2)]);
  await fs.writeFile(
    `${root}/browser-initial-diagnostic.json`,
    JSON.stringify(
      {
        initial,
        dom: await paragraphs.nth(1).evaluate((p) => ({
          html: p.outerHTML,
          width: p.getBoundingClientRect().width,
          font: getComputedStyle(p).font,
          spanStyles: [...p.querySelectorAll('span')].map((s) => ({
            font: getComputedStyle(s).font,
            size: getComputedStyle(s).fontSize,
            weight: getComputedStyle(s).fontWeight,
            family: getComputedStyle(s).fontFamily,
            spacing: getComputedStyle(s).letterSpacing,
            width: s.getBoundingClientRect().width,
          })),
          canvas: (() => {
            const c = document.createElement('canvas').getContext('2d')!;
            c.font = '10pt Arial';
            return {
              width: c.measureText('word1 word2 word3 word4 word5 word6 word7').width,
              font: c.font,
            };
          })(),
          whiteSpace: getComputedStyle(p).whiteSpace,
          wordSpacing: getComputedStyle(p).wordSpacing,
          letterSpacing: getComputedStyle(p).letterSpacing,
          textIndent: getComputedStyle(p).textIndent,
        })),
      },
      null,
      2,
    ),
  );
  const check = async (index: number, advance: number, shift: number) => {
    // Chromium resolves each zoomed line to 1/64 CSSpx. Bound that per-line
    // layout quantization instead of imposing a fixed whole-paragraph error.
    await expect
      .poll(async () => {
        const observed = await measure(index);
        return (
          (Math.abs(observed.height / nativeLines[index - 1].length - advance) * observed.scale) /
          0.75
        );
      })
      .toBeLessThan(1 / 64);
    const observed = await measure(index);
    expect(observed.lines.map(({ from, text }) => ({ from, text }))).toEqual(
      nativeLines[index - 1],
    );
    for (let i = 0; i < observed.lines.length; i++) {
      await expect
        .poll(async () =>
          Math.abs(
            (await measure(index)).lines[i].top -
              initial[index - 1].lines[0].top -
              shift -
              i * advance,
          ),
        )
        .toBeLessThan(0.15);
    }
  };
  const exports: Record<string, string> = {};
  const stages: Record<string, Awaited<ReturnType<typeof measure>>[]> = {};
  let previous = { advance: 17.25, shift: 0 };
  for (const stage of [
    { name: 'double', value: '2', advance: 23, shift: 0 },
    { name: 'minimum', value: 'custom', advance: 18, shift: 6.5 },
    { name: 'single', value: '1', advance: 11.5, shift: 0 },
  ]) {
    for (const index of [1, 2]) {
      await paragraphs.nth(index).click();
      await page
        .getByRole('combobox', { name: 'Line spacing', exact: true })
        .selectOption(stage.value);
      if (stage.value === 'custom') {
        const dialog = page.getByRole('dialog', { name: 'Line spacing', exact: true });
        await dialog
          .getByRole('combobox', { name: 'Spacing rule', exact: true })
          .selectOption('atLeast');
        const amount = dialog.getByRole('spinbutton', { name: 'Spacing amount', exact: true });
        await amount.fill('0');
        await expect(dialog.getByRole('button', { name: 'Apply', exact: true })).toBeDisabled();
        await amount.fill('18');
        await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
      }
      await check(index, stage.advance, stage.shift);
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await check(index, previous.advance, previous.shift);
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await check(index, stage.advance, stage.shift);
    }
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    for (const index of [1, 2]) await check(index, stage.advance, stage.shift);
    stages[stage.name] = await Promise.all([measure(1), measure(2)]);
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/browser-${stage.name}.docx`, bytes);
    exports[stage.name] = hash(bytes);
    await upload(`Wrapped ${stage.name}`, bytes);
    for (const index of [1, 2]) await check(index, stage.advance, stage.shift);
    previous = stage;
  }
  for (let i = 0; i < 5; i++)
    await page.getByRole('button', { name: 'Zoom out', exact: true }).click();
  for (const index of [1, 2]) await check(index, 11.5, 0);
  // Real hit testing and native caret ownership on the second hard-break line.
  const hit = await paragraphs.nth(2).evaluate((p) => {
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      if (!walker.currentNode.textContent?.includes('Second')) continue;
      const r = document.createRange();
      r.setStart(walker.currentNode, 0);
      r.setEnd(walker.currentNode, 1);
      const b = r.getBoundingClientRect();
      return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
    }
    throw new Error('Missing second line');
  });
  await page.mouse.click(hit.x, hit.y);
  await page.keyboard.press('Home');
  await page.keyboard.insertText('X');
  await expect(paragraphs.nth(2)).toContainText('XSecond line');
  await page.keyboard.press('Control+z');
  await expect(paragraphs.nth(2)).not.toContainText('XSecond');
  await check(2, 11.5, 0);
  for (let i = 0; i < 5; i++)
    await page.getByRole('button', { name: 'Zoom in', exact: true }).click();
  for (const index of [1, 2]) await check(index, 11.5, 0);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const saved = await page.evaluate(async () => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('noffice-workspace');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    try {
      return await new Promise<string>((resolve, reject) => {
        const request = db.transaction('files').objectStore('files').getAll();
        request.onsuccess = () => resolve(JSON.stringify(request.result));
        request.onerror = () => reject(request.error);
      });
    } finally {
      db.close();
    }
  });
  expect(saved).not.toContain('data-word-uniform-run-leading');
  expect(saved).not.toContain('--word-uniform-natural');
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify({ passed: true, sourceSha256: hash(source), initial, stages, exports }, null, 2),
  );
});
