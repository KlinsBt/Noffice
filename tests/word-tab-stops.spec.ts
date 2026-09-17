import { test, expect, type Locator } from '@playwright/test';
import fs from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
const reference = JSON.parse(readFileSync('tests/fixtures/native-word-tab-stop-glyphs.json', 'utf8')) as typeof import('./fixtures/native-word-tab-stop-glyphs.json');

async function characters(paragraph: Locator) {
  return paragraph.evaluate((p) => {
    const page = document.querySelector('.section-page') || document.querySelector('[aria-label="Document text"]')!;
    const bounds = page.getBoundingClientRect();
    const scale = parseFloat(getComputedStyle(p.closest('.paper-wrap')!).zoom) || 1;
    const result: { text: string; x: number; top: number }[] = [];
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const text = walker.currentNode as Text;
      if (text.parentElement?.closest('.ProseMirror-widget,[data-word-tab]')) continue;
      for (let i = 0; i < text.length; i++) {
        const range = document.createRange(); range.setStart(text, i); range.setEnd(text, i + 1);
        const rect = range.getBoundingClientRect();
        result.push({ text: text.data[i], x: (rect.left - bounds.left) * .75 / scale, top: rect.top });
      }
    }
    return result;
  });
}

test('native tab anchors survive import, real Tab editing, history and reload', async ({ page }) => {
  test.setTimeout(120000);
  const root = '.local/word-tab-stops-v2/browser/body'; await fs.mkdir(root, { recursive: true });
  const errors: string[] = []; page.on('pageerror', (e) => errors.push(e.message));
  await page.setViewportSize({ width: 1500, height: 1200 }); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles('tests/fixtures/word-tab-stops-body.docx');
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(body).toBeVisible();
  // Indented paragraph pagination is a remaining dependency. Verify the actual
  // horizontal anchors in its existing continuous view as well as paged views.
  const rows = reference.rows.filter((r) => r.document === 'body' && !['dots', 'dashes', 'line', 'heavy', 'middle-dot', 'bar'].includes(r.name));
  const reports: unknown[] = [];
  for (const row of rows) {
    const p = body.locator('p').nth(row.paragraph);
    await expect(p.locator('[data-word-tab-measured=true]')).toHaveCount(row.text.split('\t').length - 1);
    const indices: number[] = []; let at = 0;
    for (const field of row.text.split('\t')) { if (field) indices.push(at); at += field.length; }
    if (row.name === 'decimal-comma') indices.push(row.text.replaceAll('\t', '').indexOf(','));
    await expect.poll(async () => {
      const actual = await characters(p);
      if (actual.map((g) => g.text).join('') !== row.text.replaceAll('\t', '')) return 10000;
      return Math.max(...indices.map((i) => Math.abs(actual[i].x - row.glyphs[i].x)));
    }, { message: row.name }).toBeLessThanOrEqual(.15);
    reports.push({ name: row.name, actual: await characters(p), expected: row.glyphs });
  }
  await body.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.press('End');
  await page.keyboard.press('Tab'); await page.keyboard.insertText('D');
  await expect(body.locator('p').first()).toHaveText('A\tB\tC\tD');
  await page.keyboard.press('Control+z'); await page.keyboard.press('Control+z');
  await expect(body.locator('p').first()).toHaveText('A\tB\tC');
  await page.keyboard.press('Control+y'); await page.keyboard.press('Control+y');
  await expect(body.locator('p').first()).toHaveText('A\tB\tC\tD');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'word-tab-stops-body', exact: true }).click();
  await expect(body.locator('p').first()).toHaveText('A\tB\tC\tD');
  await expect(body.locator('p').first().locator('[data-word-tab-measured=true]')).toHaveCount(3);
  const download = async (name: string, label: string | RegExp) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: typeof label === 'string' }).click();
    await (await waiting).saveAs(root + '/' + name);
  };
  await download('edited.noffice', /^Noffice backup/);
  const backup = JSON.parse(await fs.readFile(root + '/edited.noffice', 'utf8'));
  expect(backup.content.tabStopsVersion).toBe(1);
  expect(backup.content.html).not.toContain('data-word-tab-measured');
  await download('edited.docx', 'DOCX file Editable in Microsoft Word');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: 'Returned tabs.docx',
    buffer: await fs.readFile(root + '/edited.docx'),
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect(body.locator('p').first()).toHaveText('A\tB\tC\tD');
  await download('reimported.docx', 'DOCX file Editable in Microsoft Word');
  expect(await fs.readFile(root + '/reimported.docx')).toEqual(await fs.readFile(root + '/edited.docx'));
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ rows: reports, errors,
    sourceHash: createHash('sha256').update(await fs.readFile('tests/fixtures/word-tab-stops-body.docx')).digest('hex'),
    exportHash: createHash('sha256').update(await fs.readFile(root + '/edited.docx')).digest('hex'),
    buildHash: await wordStoryBuildHash() }, null, 2));
});

test('native header and footer tab anchors remain on their measured lines', async ({ page }) => {
  test.setTimeout(60000);
  const root = '.local/word-tab-stops-v2/browser/stories'; await fs.mkdir(root, { recursive: true });
  const errors: string[] = []; page.on('pageerror', (e) => errors.push(e.message));
  await page.setViewportSize({ width: 1500, height: 1200 }); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles('tests/fixtures/word-tab-stops-stories.docx');
  await page.context().grantPermissions(['local-fonts']);
  await expect(page.getByRole('textbox', { name: 'Document text', exact: true })).toBeVisible();
  try {
    await expect(page.locator('.section-page')).toHaveCount(3);
    await expect(page.locator('.word-page-story')).toHaveCount(6);
    const reports: unknown[] = [];
    for (const [i, row] of reference.rows.filter((r) => r.document === 'stories').entries()) {
      const story = page.locator('.word-page-story').nth(i);
      const p = story.locator('p');
      await expect(p).toHaveCount(1);
      const actual = await characters(p);
      expect(actual.map((g) => g.text).join('')).toBe('LeftCenterRight');
      expect(new Set(actual.map((g) => Math.round(g.top))).size).toBe(1);
      const indices = [0, 4, 10];
      expect(Math.max(...indices.map((n) => Math.abs(actual[n].x - row.glyphs[n].x)))).toBeLessThanOrEqual(.15);
      reports.push({ name: row.name, actual, expected: row.glyphs });
    }
    const outputs: Record<string, string> = {};
    const download = async (name: string, label: string) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const waiting = page.waitForEvent('download');
      await page.getByRole('button', { name: label, exact: true }).click();
      await (await waiting).saveAs(root + '/' + name);
      outputs[name] = createHash('sha256').update(await fs.readFile(root + '/' + name)).digest('hex');
    };
    await download('source.docx', 'DOCX file Editable in Microsoft Word');
    expect(await fs.readFile(root + '/source.docx')).toEqual(await fs.readFile('tests/fixtures/word-tab-stops-stories.docx'));
    await download('source.pdf', 'PDF file');
    const openHeader = async () => {
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
      return page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    };
    const story = await openHeader(); await story.focus(); await page.keyboard.press('Home');
    for (let i = 0; i < 4; i++) await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.insertText('Edited');
    await expect(story).toHaveText('Edited\tCenter\tRight');
    await page.keyboard.press('Control+z'); await expect(story).toHaveText('Left\tCenter\tRight');
    await page.keyboard.press('Control+y'); await expect(story).toHaveText('Edited\tCenter\tRight');
    await page.keyboard.press('Home');
    for (let i = 0; i < 6; i++) await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Delete'); await expect(story).toHaveText('EditedCenter\tRight');
    await page.keyboard.press('Tab'); await expect(story).toHaveText('Edited\tCenter\tRight');
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    const painted = page.locator('.word-page-story').filter({ hasText: 'Edited' });
    await expect(painted).toHaveCount(3);
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(painted).toHaveCount(0);
    await page.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(painted).toHaveCount(3);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: 'word-tab-stops-stories', exact: true }).click();
    await expect(painted).toHaveCount(3);
    await download('edited.docx', 'DOCX file Editable in Microsoft Word');
    await download('edited.pdf', 'PDF file');
    await page.locator('input[type=file][multiple]').setInputFiles({ name: 'Returned stories.docx',
      buffer: await fs.readFile(root + '/edited.docx'),
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    await expect(await openHeader()).toHaveText('Edited\tCenter\tRight');
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await download('reimported.docx', 'DOCX file Editable in Microsoft Word');
    expect(outputs['reimported.docx']).toBe(outputs['edited.docx']);
    await download('reimported.pdf', 'PDF file');
    expect(errors).toEqual([]);
    await fs.writeFile(root + '/browser-report.json', JSON.stringify({ rows: reports, errors,
      sourceHash: createHash('sha256').update(await fs.readFile('tests/fixtures/word-tab-stops-stories.docx')).digest('hex'),
      outputs, buildHash: await wordStoryBuildHash() }, null, 2));
  } catch (error) {
    await fs.writeFile(root + '/stability.json', JSON.stringify(await page.evaluate(async () => {
      const views = [...document.querySelectorAll('.word-story-measure-host .ProseMirror')].map((e) =>
        (e as unknown as { editor: import('@tiptap/core').Editor }).editor.view);
      const results = []; let previous: unknown[] = [];
      for (let i = 0; i < 20; i++) {
        await new Promise(requestAnimationFrame);
        const states = views.flatMap((v) => v.state.plugins.filter((p) => /wordTabLayout|wordFontLineMetrics/.test((p as unknown as { key: string }).key)).map((p) => p.getState(v.state)));
        results.push(states.map((s, i) => ({ same: s === previous[i], signature: s?.signature })));
        previous = states;
      }
      return results;
    }), null, 2));
    await fs.writeFile(root + '/layout-failure.json', JSON.stringify(await page.evaluate(() => ({
      body: document.querySelector('[aria-label="Document text"]')?.outerHTML,
      hosts: [...document.querySelectorAll('.word-story-measure-host')].map((h) => h.outerHTML),
      paragraphs: [...document.querySelectorAll('.word-story-measure-host p')].map((p) => {
        const view = (p.closest('.ProseMirror') as unknown as { editor: import('@tiptap/core').Editor }).editor.view;
        const chars = [];
        for (let i = 1; i < view.state.doc.content.size - 1; i++) {
          const a = view.domAtPos(i, 1), b = view.domAtPos(i + 1, -1), r = document.createRange();
          r.setStart(a.node, a.offset); r.setEnd(b.node, b.offset);
          chars.push({ i, text: r.toString(), rects: [...r.getClientRects()].map((r) => r.toJSON()) });
        }
        const s = getComputedStyle(p); return { html: p.innerHTML, height: s.height, line: s.lineHeight,
          chars,
          before: s.marginTop, after: s.marginBottom, bounds: p.getBoundingClientRect().toJSON(),
          children: [...p.querySelectorAll('span')].map((e) => ({ text: e.textContent,
            rect: e.getBoundingClientRect().toJSON(), font: getComputedStyle(e).fontSize,
            line: getComputedStyle(e).lineHeight, align: getComputedStyle(e).verticalAlign })) };
      }),
      stories: [...document.querySelectorAll('.word-page-story')].map((h) => h.outerHTML),
      errors: document.body.innerText,
    })), null, 2));
    throw error;
  }
});

test('native center/right/decimal fields respect the text margin without changing saved stops', async ({ page }) => {
  const native = JSON.parse(readFileSync('tests/fixtures/native-word-tab-boundaries.json', 'utf8')) as typeof import('./fixtures/native-word-tab-boundaries.json');
  const root = '.local/word-tab-boundaries/browser'; await fs.mkdir(root, { recursive: true });
  await page.setViewportSize({ width: 1500, height: 1200 }); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles('tests/fixtures/word-tab-boundaries.docx');
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(body).toBeVisible();
  const reports = [];
  for (const row of native.rows.filter((r) => r.alignment !== 0)) {
    const p = body.locator('p').nth(row.paragraph);
    await expect(p.locator('[data-word-tab-measured=true]')).toHaveCount(1);
    await expect.poll(async () => {
      const actual = await characters(p);
      return actual.length === 6 ? Math.abs(actual[1].x - row.glyphs[1].x) : 10000;
    }, { message: JSON.stringify([row.alignment, row.position]) }).toBeLessThanOrEqual(.15);
    const actual = await characters(p);
    expect(new Set(actual.map((g) => Math.round(g.top))).size).toBe(1);
    expect(await p.getAttribute('data-word-tabs')).toContain('"position":' + Math.round(row.position * 20));
    reports.push({ alignment: row.alignment, position: row.position, actual, expected: row.glyphs });
  }
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ rows: reports,
    sourceHash: createHash('sha256').update(await fs.readFile('tests/fixtures/word-tab-boundaries.docx')).digest('hex'),
    buildHash: await wordStoryBuildHash() }, null, 2));
});
