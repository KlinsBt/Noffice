import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import native from './fixtures/native-word-tab-leaders.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

test('leader cells retain physical origins through zoom, pointer editing and actual PDFs', async ({ page }) => {
  test.setTimeout(90000);
  const run = process.env.NOFFICE_LEADER_ZOOM_RUN || 'zoom-v1';
  expect(run).toMatch(/^zoom-v\d+$/);
  const root = '.local/word-tab-leader-render/' + run;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-tab-leaders.docx');
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const errors: string[] = [], outputs: Record<string, string> = {};
  page.on('pageerror', (error) => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.setViewportSize({ width: 1500, height: 1200 });await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: 'leader-zoom.docx', buffer: source,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  const first = body.locator('p').first();
  let zoom = 100;
  const setZoom = async (value: number) => {
    while (zoom !== value) {
      const direction = value > zoom ? 10 : -10;
      await page.getByRole('button', { name: direction > 0 ? 'Zoom in' : 'Zoom out', exact: true }).click();
      zoom += direction;
    }
    await expect(body.locator('[data-word-tab-leader-painted=true]')).toHaveCount(90);
  };
  const measure = () => body.evaluate((host) => {
    const scale = parseFloat(getComputedStyle(host.closest('.paper-wrap')!).zoom) || 1;
    return [...host.querySelectorAll('p')].map((p) => {
      const semantic = p.querySelector<HTMLElement>('[data-word-tab]');
      if (!semantic) return null;
      const tab = semantic.closest<HTMLElement>('[data-word-tab-measured]') || semantic;
      const offsets = JSON.parse(tab.dataset.wordTabLeaderOffsets || 'null') as number[] | null;
      const x = (tab.getBoundingClientRect().left - host.getBoundingClientRect().left) * .75 / scale;
      const pseudo = getComputedStyle(tab, '::after');
      return { x, offsets, color: pseudo.color, glyph: pseudo.content,
        points: parseFloat(pseudo.fontSize) * .75, scale };
    });
  });
  const recorded: unknown[] = [];
  const check = async (edited = false) => {
    await expect.poll(async () => {
      const actual = await measure();
      if (actual.length !== native.cases.length) return 10000;
      return Math.max(0, ...actual.flatMap((row, i) => {
        const expected = edited && i === 0 ? native.cases[2] : native.cases[i];
        if (expected.noTab) return row === null ? [0] : [10000];
        if (!row?.offsets || row.offsets.length !== expected.leaders.length) return [10000];
        return row.offsets.map((offset, j) => Math.abs(row.x + offset * .75 - expected.leaders[j].x));
      }));
    }).toBeLessThanOrEqual(.15);
    const actual = await measure();
    for (const [i, row] of actual.entries()) {
      if (!row) continue;
      const expected = edited && i === 0 ? native.cases[2] : native.cases[i];
      expect(row.scale).toBeCloseTo(zoom / 100, 5);
      expect(row.glyph).toBe(JSON.stringify(expected.leaders[0]?.text || ''));
      if (expected.leaders.length) {
        expect(row.color).toBe(`rgb(${expected.leaders[0].color.slice(0, 3).join(', ')})`);
        expect(Math.abs(row.points - expected.leaders[0].size)).toBeLessThanOrEqual(.15);
      }
    }
    recorded.push({ zoom, edited, actual });
  };
  const download = async (name: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await pending).saveAs(root + '/' + name);
    outputs[name] = hash(await fs.readFile(root + '/' + name));
  };
  await check();
  for (const value of [50, 80, 150, 100]) {
    await setZoom(value);await check();await download(`source-${value}.pdf`);
  }
  await setZoom(50);await first.scrollIntoViewIfNeeded();
  const hit = await first.evaluate((p) => {
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const text = walker.currentNode as Text;
      if (text.data !== 'A') continue;
      const range = document.createRange();range.selectNodeContents(text);
      const rect = range.getBoundingClientRect();
      return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
    }
    throw Error('Missing first authored glyph');
  });
  await page.mouse.click(hit.x, hit.y);await page.keyboard.press('Home');
  await page.keyboard.press('Shift+ArrowRight');await page.keyboard.insertText('Edited');
  await expect(first).toHaveText('Edited\tB');await check(true);
  await page.keyboard.press('Control+z');await expect(first).toHaveText('A\tB');await check();
  await page.keyboard.press('Control+y');await expect(first).toHaveText('Edited\tB');await check(true);
  await download('edited-50.pdf');await download('edited.docx', 'DOCX file Editable in Microsoft Word');
  await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'leader-zoom', exact: true }).click();zoom = 100;
  await check(true);await download('reloaded.pdf');
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(source), outputs, recorded, errors,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-tab-leader-zoom.spec.ts')) }, null, 2));
});
