import { test, expect, type Locator } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import native from './fixtures/native-word-tab-leaders.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

test('native leader cells paint through real editing, history, reload and DOCX export', async ({ page }) => {
  test.setTimeout(90000);const root = '.local/word-tab-leader-render/browser';await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-tab-leaders.docx');
  const errors: string[] = [];page.on('pageerror', (error) => errors.push(error.message));
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const outputs: Record<string, string> = {};
  const download = async (name: string, label: string | RegExp) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: typeof label === 'string' }).click();await (await pending).saveAs(root + '/' + name);
    const bytes = await fs.readFile(root + '/' + name);outputs[name] = hash(bytes);return bytes;
  };
  await page.setViewportSize({ width: 1500, height: 1200 });await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles('tests/fixtures/word-tab-leaders.docx');
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(page.locator('.section-page')).toHaveCount(6);
  const paint = (p: Locator) => p.locator('[data-word-tab]').evaluate((semanticTab) => {
    const tab = semanticTab.closest('[data-word-tab-measured]') || semanticTab;
    const host = tab.closest('[aria-label="Document text"]')!, canvas = tab.closest('.paper-wrap')!;
    const scale = parseFloat(getComputedStyle(canvas).zoom) || 1;
    const pseudo = getComputedStyle(tab, '::after');
    const offsets = JSON.parse(tab.getAttribute('data-word-tab-leader-offsets') || 'null') as number[] | null;
    return { offsets, x: (tab.getBoundingClientRect().left - host.getBoundingClientRect().left) * .75 / scale,
      content: pseudo.content, shadows: (pseudo.textShadow.match(/rgba?\(/g) || []).length,
      color: pseudo.color, size: parseFloat(pseudo.fontSize) * .75 };
  });
  const matches = async (paragraph: Locator, row: typeof native.cases[number]) => {
    await expect(paragraph.locator('[data-word-tab-leader-painted=true]')).toHaveCount(1);
    await expect.poll(async () => {
      const actual = await paint(paragraph);
      if (actual.offsets?.length !== row.leaders.length) return 10000;
      return Math.max(0, ...actual.offsets.map((x, i) => Math.abs(actual.x + x * .75 - row.leaders[i].x)));
    }, { message: row.name }).toBeLessThanOrEqual(.15);
    const actual = await paint(paragraph);
    expect(actual.shadows).toBe(Math.max(0, row.leaders.length - 1));
    expect(actual.content).toBe(JSON.stringify(row.leaders[0]?.text || ''));
    if (row.leaders.length) {
      expect(actual.color).toBe(`rgb(${row.leaders[0].color.slice(0, 3).join(', ')})`);
      expect(Math.abs(actual.size - row.leaders[0].size)).toBeLessThanOrEqual(.15);
    }
    return actual;
  };
  const rows = [];
  for (const [index, row] of native.cases.entries()) {
    const paragraph = body.locator('p').nth(index);
    if (row.noTab) { await expect(paragraph.locator('[data-word-tab]')).toHaveCount(0);continue; }
    rows.push({ name: row.name, actual: await matches(paragraph, row) });
  }
  for (let group = 0; group < 5; group++) {
    await body.locator('p').nth(group * 19).scrollIntoViewIfNeeded();
    await page.screenshot({ path: root + `/source-group-${group + 1}.png` });
  }
  await page.emulateMedia({ media: 'print' });
  await page.pdf({ path: root + '/source-css.pdf', preferCSSPageSize: true, printBackground: true });
  await page.emulateMedia({ media: 'screen' });
  outputs['source-css.pdf'] = hash(await fs.readFile(root + '/source-css.pdf'));
  const originalText = await body.innerText(), rejectedDownloads: string[] = [];
  const rejectedDownload = (file: { suggestedFilename(): string }) => rejectedDownloads.push(file.suggestedFilename());
  page.on('download', rejectedDownload);
  await page.context().grantPermissions([]);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('Font access was not granted');
  expect(await body.innerText()).toBe(originalText);expect(rejectedDownloads).toEqual([]);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();page.off('download', rejectedDownload);
  await page.context().grantPermissions(['local-fonts']);
  await page.evaluate(() => {
    const target = window as any;target.leaderOriginalFonts = target.queryLocalFonts;
    target.queryLocalFonts = async () => (await target.leaderOriginalFonts.call(window))
      .filter((font: { family: string; style: string }) => !(font.family === 'Arial' && /^bold$/i.test(font.style)));
  });
  page.on('download', rejectedDownload);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('The bold font Arial is unavailable');
  expect(await body.innerText()).toBe(originalText);expect(rejectedDownloads).toEqual([]);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();page.off('download', rejectedDownload);
  await page.evaluate(() => { const target = window as any;target.queryLocalFonts = target.leaderOriginalFonts; });
  await download('source.pdf', 'PDF file');
  // A wrong face returned by the local-font API must be caught by the independent
  // exported-font comparison, even when every measured origin is unchanged.
  await page.evaluate(() => {
    const target = window as any;target.queryLocalFonts = async () => {
      const fonts = await target.leaderOriginalFonts.call(window);
      const regular = fonts.find((font: { family: string; style: string }) => font.family === 'Arial' && /^regular$/i.test(font.style));
      if (!regular) throw Error('Missing control font');
      return fonts.map((font: { family: string; style: string }) => font.family === 'Arial' && /^bold$/i.test(font.style)
        ? { family: 'Arial', style: 'Bold', blob: () => regular.blob() } : font);
    };
  });
  await download('wrong-bold-control.pdf', 'PDF file');
  await page.evaluate(() => { const target = window as any;target.queryLocalFonts = target.leaderOriginalFonts;delete target.leaderOriginalFonts; });
  await body.focus();await page.keyboard.press('Control+Home');await page.keyboard.press('Shift+ArrowRight');await page.keyboard.press('Control+b');
  await expect(body.locator('p').first().locator('strong')).toHaveText('A');
  page.on('download', rejectedDownload);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('needs measured leader glyphs and baselines');
  expect(rejectedDownloads).toEqual([]);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();page.off('download', rejectedDownload);
  await body.focus();await page.keyboard.press('Control+z');await expect(body.locator('p').first().locator('strong')).toHaveCount(0);
  await page.keyboard.press('Control+y');await expect(body.locator('p').first().locator('strong')).toHaveText('A');
  await page.keyboard.press('Control+z');await expect(body.locator('p').first().locator('strong')).toHaveCount(0);
  await matches(body.locator('p').first(), native.cases[0]);
  await download('source-recovered.pdf', 'PDF file');
  await body.focus();await page.keyboard.press('Control+Home');await page.keyboard.press('Shift+ArrowRight');await page.keyboard.insertText('Edited');
  await expect(body.locator('p').first()).toHaveText('Edited\tB');
  await matches(body.locator('p').first(), native.cases.find((row) => row.name === 'dot-edited')!);
  await page.keyboard.press('Control+z');await matches(body.locator('p').first(), native.cases[0]);
  await page.keyboard.press('Control+y');await expect(body.locator('p').first()).toHaveText('Edited\tB');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const backup = JSON.parse((await download('edited.noffice', /^Noffice backup/)).toString());
  expect(Buffer.from(backup.original.base64, 'base64')).toEqual(source);
  expect(backup.content.html).not.toContain('data-word-tab-leader-painted');
  await download('edited.docx', 'DOCX file Editable in Microsoft Word');
  await download('edited.pdf', 'PDF file');
  await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'word-tab-leaders', exact: true }).click();
  await matches(body.locator('p').first(), native.cases.find((row) => row.name === 'dot-edited')!);
  await download('reloaded.pdf', 'PDF file');
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(source), rows, outputs, errors,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-tab-leaders.spec.ts')) }, null, 2));
});
