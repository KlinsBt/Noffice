import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';

test('mixed-size Word lines match native advances through editing, history and exports', async ({
  page,
}) => {
  const root = '.local/word-mixed-lines';
  const source = await fs.readFile('tests/fixtures/word-mixed-lines.docx');
  const oracle = JSON.parse(
    await fs
      .readFile('tests/fixtures/native-word-mixed-lines.json', 'utf8')
      .then((s) => s.replace(/^\uFEFF/, '')),
  );
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  expect(hash(source)).toBe(oracle.sourceSha256);
  await fs.mkdir(root, { recursive: true });
  await fs.writeFile(`${root}/source.docx`, source);
  await page.setViewportSize({ width: 1500, height: 1200 });
  await page.goto('/');
  const upload = async (name: string, buffer: Buffer) => {
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({
        name: `${name}.docx`,
        buffer,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  await upload('Mixed lines', source);
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const paragraphs = editor.locator('p');
  await expect(paragraphs).toHaveCount(7);
  const heights: number[] = [];
  for (let i = 0; i < 6; i++) {
    const expected = oracle.source[i + 1].chars[0].y - oracle.source[i].chars[0].y;
    // Word COM positions quantize to twips/printer units; allow 0.15pt, not whole pixels.
    await expect
      .poll(async () => Math.abs((await paragraphs.nth(i).boundingBox())!.height * 0.75 - expected))
      .toBeLessThan(0.15);
    await expect
      .poll(async () =>
        Math.abs(
          ((await paragraphs.nth(i + 1).boundingBox())!.y -
            (await paragraphs.nth(i).boundingBox())!.y) *
            0.75 -
            expected,
        ),
      )
      .toBeLessThan(0.15);
    heights.push((await paragraphs.nth(i).boundingBox())!.height);
  }
  const baselines = await paragraphs.evaluateAll((ps) =>
    ps.slice(0, 6).map((p) => {
      const marker = document.createElement('span');
      marker.style.cssText =
        'display:inline-block;width:0;height:0;padding:0;line-height:0;vertical-align:baseline';
      p.append(marker);
      const baseline = (marker.getBoundingClientRect().top - p.getBoundingClientRect().top) * 0.75;
      marker.remove();
      return baseline;
    }),
  );
  await fs.writeFile(
    `${root}/browser-source-metrics.json`,
    JSON.stringify({ heights, baselines }, null, 2),
  );
  await page.screenshot({ path: `${root}/browser-source.png`, fullPage: true });
  await editor.evaluate((el) => (el.style.width = '90px'));
  await expect(paragraphs.nth(1)).not.toHaveAttribute('data-word-mixed-leading');
  await editor.evaluate((el) => el.style.removeProperty('width'));
  await expect(paragraphs.nth(1)).toHaveAttribute('data-word-mixed-leading', 'true');
  // Word adds a hidden _GoBack bookmark to paragraph 1; its formatting remains guarded.
  const first = paragraphs.nth(1);
  await first.click();
  await page.keyboard.press('Home');
  for (let i = 0; i < 6; i++) await page.keyboard.press('ArrowRight');
  for (let i = 0; i < 5; i++) await page.keyboard.press('Shift+ArrowRight');
  expect(await page.evaluate(() => getSelection()?.toString())).toBe('Large');
  const size = page.getByRole('spinbutton', { name: 'Font size', exact: true });
  await size.fill('20');
  await size.press('Enter');
  await expect.poll(async () => (await first.boundingBox())!.height * 0.75).toBeCloseTo(34.5, 1);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(async () => (await first.boundingBox())!.height * 0.75).toBeCloseTo(51.75, 1);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect.poll(async () => (await first.boundingBox())!.height * 0.75).toBeCloseTo(34.5, 1);
  const exports: Record<string, string> = {};
  const saveAndExport = async (stage: string) => {
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    await expect(first).toHaveText('Small Large end');
    await expect
      .poll(async () => (await first.boundingBox())!.height * 0.75)
      .toBeCloseTo(stage === 'size' ? 34.5 : 50, 1);
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/browser-${stage}.docx`, bytes);
    exports[stage] = hash(bytes);
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    expect(xml).not.toContain('measured');
    await upload(`Mixed ${stage}`, bytes);
    await expect
      .poll(async () => (await first.boundingBox())!.height * 0.75)
      .toBeCloseTo(stage === 'size' ? 34.5 : 50, 1);
  };
  await saveAndExport('size');
  await first.click();
  await page.keyboard.press('Home');
  await page.getByRole('combobox', { name: 'Line spacing', exact: true }).selectOption('custom');
  const dialog = page.getByRole('dialog', { name: 'Line spacing', exact: true });
  await dialog.getByRole('combobox', { name: 'Spacing rule', exact: true }).selectOption('atLeast');
  const amount = dialog.getByRole('spinbutton', { name: 'Spacing amount', exact: true });
  await amount.fill('0');
  await expect(dialog.getByRole('button', { name: 'Apply', exact: true })).toBeDisabled();
  await amount.fill('50');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await fs.writeFile(
    `${root}/minimum-dom.json`,
    JSON.stringify(
      await first.evaluate((p) => ({
        html: p.outerHTML,
        height: p.getBoundingClientRect().height,
        style: {
          line: getComputedStyle(p).lineHeight,
          before: getComputedStyle(p).paddingTop,
          after: getComputedStyle(p).paddingBottom,
          min: getComputedStyle(p).minHeight,
        },
      })),
      null,
      2,
    ),
  );
  await expect.poll(async () => (await first.boundingBox())!.height * 0.75).toBeCloseTo(50, 1);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(async () => (await first.boundingBox())!.height * 0.75).toBeCloseTo(34.5, 1);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await saveAndExport('minimum');
  // Measurements must not enter persisted content or history.
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const stored = await page.evaluate(async () => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const r = indexedDB.open('noffice-workspace');
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
    const records = await new Promise<unknown[]>((resolve, reject) => {
      const r = db.transaction('files').objectStore('files').getAll();
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    });
    db.close();
    return JSON.stringify(records);
  });
  expect(stored).toContain('Mixed minimum');
  expect(stored).not.toContain('data-word-measured-leading');
  expect(stored).not.toContain('--word-mixed-');
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      { passed: true, sourceSha256: hash(source), heights, baselines, exports },
      null,
      2,
    ),
  );
});
