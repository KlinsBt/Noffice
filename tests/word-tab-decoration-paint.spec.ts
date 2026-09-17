import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import native from './fixtures/word-pdf-tab-decorations/native-paint.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

for (const [mode, original] of [['underline', 1], ['strike', 2], ['both', 3]] as const)
  test(`visible ${mode} tab paint survives history, reload and 50% zoom`, async ({ page }) => {
    test.setTimeout(90000);
    const run = process.env.NOFFICE_TAB_PAINT_RUN || 'paint-browser-v1';
    if (!/^paint-browser-v\d+$/.test(run)) throw Error('Invalid tab paint run');
    const root = `.local/word-pdf-tab-decorations/${run}/${mode}`; await fs.mkdir(root, { recursive: true });
    await page.setViewportSize({ width: 1500, height: 1200 }); await page.goto('/');
    const source = `tests/fixtures/word-pdf-tab-decorations/body-${mode}.docx`;
    await page.locator('input[type=file][multiple]').setInputFiles(source);
    const body = page.getByRole('textbox', { name: 'Document text', exact: true }); await waitWordLayout(body);
    const text = await body.textContent(), rows: unknown[] = [];
    const capture = async (stage: string, marks: number) => {
      await expect.poll(() => body.locator('[data-word-tab]').first().evaluate(tab =>
        Number((tab.closest('[data-word-tab-measured]') || tab).getAttribute('data-word-tab-decoration') || 0))).toBe(marks);
      for (const zoom of [1, .5]) {
        if (zoom === .5) for (let i = 0; i < 5; i++) await page.getByRole('button', { name: 'Zoom out', exact: true }).click();
        await waitWordLayout(body);
        const png = await body.screenshot({ path: `${root}/${stage}-${zoom}.png` });
        const pixels = await page.evaluate(async ({ png, paints, zoom }) => {
          const image = new Image(); image.src = 'data:image/png;base64,' + png; await image.decode();
          const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
          const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
          const data = context.getImageData(0, 0, image.width, image.height).data;
          const x = Math.round((paints[0].left + paints[0].right) / 2 / .75 * zoom);
          const top = Math.floor(Math.min(...paints.map(p => p.top)) / .75 * zoom) - 2;
          const bottom = Math.ceil(Math.max(...paints.map(p => p.bottom)) / .75 * zoom) + 2;
          const scan: { y: number; ink: number }[] = []; let bands = 0, previous = false;
          for (let y = top; y <= bottom; y++) {
            let ink = 0;
            for (let column = x - 2; column <= x + 2; column++) {
              const offset = (y * image.width + column) * 4;
              ink += Math.max(0, (255 - data[offset]) / (255 - paints[0].color[0])) / 5;
            }
            const colored = ink > .08;
            if (colored && !previous) bands++;
            previous = colored; scan.push({ y, ink });
          }
          return { bands, scan, width: image.width, height: image.height };
        }, { png: png.toString('base64'), paints: native.paints, zoom });
        rows.push({ stage, marks, zoom, pixels });
        await fs.writeFile(`${root}/pixels.json`, JSON.stringify(rows, null, 2));
        expect(pixels.bands, `${stage} at ${zoom * 100}% must retain each separate visible mark`).toBe((marks & 1 ? 1 : 0) + (marks & 2 ? 1 : 0));
        await expect.poll(() => body.textContent()).toBe(text);
        if (zoom === .5) for (let i = 0; i < 5; i++) await page.getByRole('button', { name: 'Zoom in', exact: true }).click();
      }
    };
    await capture('source', original);
    await body.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Shift+ArrowRight'); await page.keyboard.press('Control+u'); await page.keyboard.press('ArrowRight');
    await capture('edited', original ^ 1);
    await page.getByRole('button', { name: 'Undo', exact: true }).click(); await body.focus(); await page.keyboard.press('ArrowRight');
    await capture('undo', original);
    await page.getByRole('button', { name: 'Redo', exact: true }).click(); await body.focus(); await page.keyboard.press('ArrowRight');
    await capture('redo', original ^ 1);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click(); await waitWordLayout(body);
    await capture('reloaded', original ^ 1);
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    await fs.writeFile(`${root}/report.json`, JSON.stringify({ rows, buildHash: await wordStoryBuildHash(),
      sourceHash: hash(await fs.readFile(source)), testHash: hash(await fs.readFile('tests/word-tab-decoration-paint.spec.ts')),
      nativeReferenceHash: hash(await fs.readFile('tests/fixtures/word-pdf-tab-decorations/native-paint.json')),
      scope: 'Actual screenshot pixel bands detect missing visible underline/strike through editing/history/reload/zoom. This is not full pixel/geometry/native UI equivalence.' }, null, 2));
  });
