import { test, expect, type Locator } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import plain from './fixtures/word-pdf-tab-decorations/reference.json' with { type: 'json' };
import leaders from './fixtures/word-pdf-tab-leader-decorations/reference.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const samples = (['body', 'header', 'footer'] as const).flatMap(context => [
  ...plain.rows.filter(row => row.context === (context === 'body' ? 'body' : 'stories')).map(row => ({ ...row, context, kind: 'plain' })),
  ...leaders.rows.filter(row => row.context === (context === 'body' ? 'body' : 'stories') && row.mode === 'both').map(row => ({ ...row, context, kind: 'leader' })),
]);
const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
const tabMarks = (editor: Locator) => editor.locator('[data-word-tab]').first().evaluate(tab =>
  (tab.closest('u') ? 1 : 0) | (tab.closest('s,strike,del') ? 2 : 0));

for (const sample of samples) test(`tab ink ${sample.context} ${sample.name}: print, zoom and formatting history`, async ({ page }) => {
  test.setTimeout(120000);
  const run = process.env.NOFFICE_TAB_DECORATION_PRINT_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid tab print run');
  const name = `${sample.context}-${sample.name}`, root = `.local/word-tab-decoration-print/${run}/${name}`;
  await fs.mkdir(root, { recursive: true });
  expect(hash(await fs.readFile(sample.path))).toBe(sample.sourceHash);
  const outputs: Record<string, string> = {}, errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.setViewportSize({ width: 1500, height: 1200 }); await page.context().grantPermissions(['local-fonts']);
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(sample.path);
  const body = page.getByRole('textbox', { name: 'Document text', exact: true }); await waitWordLayout(body);
  let zoom = 100;
  const print = async (stage: string, target: number) => {
    while (zoom !== target) {
      const step = target > zoom ? 10 : -10;
      await page.getByRole('button', { name: step > 0 ? 'Zoom in' : 'Zoom out', exact: true }).click(); zoom += step;
    }
    await waitWordLayout(body);
    const model = () => body.evaluate(el => (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getJSON());
    const before = await model(), stories = await page.locator('.word-page-story').allTextContents();
    await page.emulateMedia({ media: 'print' });
    const file = `${stage}-${target}.pdf`; await page.pdf({ path: `${root}/${file}`, preferCSSPageSize: true, printBackground: true });
    await page.emulateMedia({ media: 'screen' });
    // Screen stories are republished by the scheduled layout frame after print.
    // Require their exact restoration, including when the print tree was removed.
    await waitWordLayout(body);
    expect(await model()).toEqual(before);
    await expect.poll(() => page.locator('.word-page-story').allTextContents()).toEqual(stories);
    outputs[file] = hash(await fs.readFile(`${root}/${file}`));
  };
  const docx = async (stage: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
    const file = `${stage}.docx`; await (await pending).saveAs(`${root}/${file}`); outputs[file] = hash(await fs.readFile(`${root}/${file}`));
  };
  for (const value of [50, 100, 150]) await print('source', value);
  await docx('source'); expect(outputs['source.docx']).toBe(sample.sourceHash);
  let editor = body;
  if (sample.context !== 'body') {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: new RegExp(`^Section 1 .*Default ${sample.context}`) }).click();
    editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  }
  const original = ({ underline: 1, strike: 2, both: 3 } as Record<string, number>)[sample.mode];
  await expect.poll(() => tabMarks(editor)).toBe(original);
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Shift+ArrowRight'); await page.keyboard.press('Control+u'); await page.keyboard.press('Control+Shift+s');
  await expect.poll(() => tabMarks(editor)).toBe(original ^ 3);
  await page.keyboard.press('Control+z'); await page.keyboard.press('Control+z');
  await expect.poll(() => tabMarks(editor)).toBe(original);
  await page.keyboard.press('Control+y'); await page.keyboard.press('Control+y');
  await expect.poll(() => tabMarks(editor)).toBe(original ^ 3);
  if (sample.context !== 'body') await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await waitWordLayout(body); await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  for (const value of [150, 100, 50]) await print('edited', value);
  await docx('edited');
  const title = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: title, exact: true }).click(); await waitWordLayout(body);
  // Zoom preference survives reload; keep the current value rather than
  // assuming that reload silently returned to100%.
  zoom = await body.evaluate(el => Math.round((parseFloat(getComputedStyle(el.closest('.paper-wrap')!).zoom) || 1) * 100));
  await print('reloaded', 100); await docx('reloaded'); expect(errors).toEqual([]);
  await fs.writeFile(`${root}/report.json`, JSON.stringify({ name, fixtureName: sample.name, context: sample.context,
    kind: sample.kind, mode: sample.mode, sourceHash: sample.sourceHash, outputs, errors, buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-tab-decoration-print.spec.ts')),
    scope: 'Actual CSS PDF at50/100/150% through tab formatting/history/reload in body/header/footer, all marks and five leader kinds. Native ink geometry, full pages and actual DOCX comparisons are separate gates.' }, null, 2));
});
