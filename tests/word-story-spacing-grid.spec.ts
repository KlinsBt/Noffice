import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import reference from './fixtures/native-word-story-spacing-grid.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word story spacing grid ${row.name}: layout, history and original-file recovery`, async ({ page }) => {
  const run = process.env.NOFFICE_STORY_GRID_RUN || 'grid-browser-v1';
  if (!/^grid-browser-v\d+$/.test(run)) throw Error('Invalid spacing evidence folder');
  const root = `.local/word-story-spacing/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile(`tests/fixtures/word-story-spacing-grid-${row.name}.docx`);
  expect(hash(source)).toBe(row.docxHash);
  const errors: string[] = []; page.on('pageerror', (error) => errors.push(error.message));
  const outputs: Record<string, string> = {}, checks: unknown[] = [];
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: `spacing-${row.name}.docx`,
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const check = async () => {
    await expect(page.locator('.section-page')).toHaveCount(2);
    await expect(page.locator('.word-page-story [data-word-tab-leader-painted=true]')).toHaveCount(20);
    for (const kind of ['header', 'footer'] as const) {
      const first = page.locator(`.word-page-story[data-word-story*="${kind}"]`).first();
      const expected = row.origins[kind];
      let positions: number[] = [];
      await expect.poll(async () => {
        positions = await first.locator('p').evaluateAll(paragraphs =>
          paragraphs.map(p => p.getBoundingClientRect().top * .75));
        return positions.length === 5 ? Math.max(...positions.map((value, index) =>
          Math.abs((value - positions[0]) - (expected[index] - expected[0])))) : Infinity;
      }).toBeLessThanOrEqual(.15);
      checks.push({ kind, positions });
    }
  };
  const download = async (name: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await pending).saveAs(root + '/' + name); outputs[name] = hash(await fs.readFile(root + '/' + name));
  };
  await check(); await download('source.pdf'); await download('original.docx', 'DOCX file Editable in Microsoft Word');
  expect(await fs.readFile(root + '/original.docx')).toEqual(source);
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${row.kind}`) }).click();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.insertText('Edited'); await expect(editor.locator('p').first()).toHaveText('Edited\tB');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText('A\tB');
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText('Edited\tB');
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click(); await check();
  const story = page.locator(`.word-page-story[data-word-story*="${row.kind}"]`).first();
  await expect(story.locator('p').first()).toHaveText('Edited\tB');
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(story.locator('p').first()).toHaveText('A\tB'); await check();
  await page.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(story.locator('p').first()).toHaveText('Edited\tB'); await check();
  await download('edited.docx', 'DOCX file Editable in Microsoft Word'); await download('edited.pdf');
  const stored = () => page.evaluate(() => new Promise<unknown[]>((resolve, reject) => {
    const request = indexedDB.open('noffice-workspace'); request.onerror = () => reject(request.error);
    request.onsuccess = () => {
      const db = request.result, tx = db.transaction('files'), all = tx.objectStore('files').getAll();
      tx.oncomplete = () => { db.close(); resolve(all.result); };
      tx.onerror = tx.onabort = () => { db.close(); reject(tx.error); };
    };
  }));
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const beforePrint = await stored();
  await page.emulateMedia({ media: 'print' });
  await page.pdf({ path: root + '/browser-print.pdf', preferCSSPageSize: true, printBackground: true });
  await page.emulateMedia({ media: 'screen' }); outputs['browser-print.pdf'] = hash(await fs.readFile(root + '/browser-print.pdf'));
  expect(await stored()).toEqual(beforePrint);
  await download('after-print.docx', 'DOCX file Editable in Microsoft Word');
  // Rebuilt ZIP entries carry a new DOS timestamp; every OOXML/asset byte must
  // remain identical, including parts unrelated to the edited story.
  const beforeZip = await JSZip.loadAsync(await fs.readFile(root + '/edited.docx'));
  const afterZip = await JSZip.loadAsync(await fs.readFile(root + '/after-print.docx'));
  expect(Object.keys(afterZip.files).sort()).toEqual(Object.keys(beforeZip.files).sort());
  for (const name of Object.keys(beforeZip.files))
    expect(await afterZip.files[name].async('uint8array'), name).toEqual(await beforeZip.files[name].async('uint8array'));
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `spacing-${row.name}`, exact: true }).click(); await check();
  await expect(story.locator('p').first()).toHaveText('Edited\tB'); await download('reloaded.pdf');
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles({ name: `grid-reimport-${row.name}.docx`, buffer: await fs.readFile(root + '/edited.docx'), mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await check(); await expect(story.locator('p').first()).toHaveText('Edited\tB'); await download('reimported.pdf'); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: row.docxHash, outputs, checks, errors,
    buildHash: await wordStoryBuildHash(), referenceHash: hash(await fs.readFile('tests/fixtures/native-word-story-spacing-grid.json')),
    testHash: hash(await fs.readFile('tests/word-story-spacing-grid.spec.ts')) }, null, 2));
});
