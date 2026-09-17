import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-story-spacing.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word story spacing ${row.name}: layout, history and original-file recovery`, async ({ page }) => {
  const run = process.env.NOFFICE_STORY_SPACING_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid spacing evidence folder');
  const root = `.local/word-story-spacing/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile(`tests/fixtures/word-story-spacing-${row.name}.docx`);
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
      const positions = await first.locator('p').evaluateAll((paragraphs) =>
        paragraphs.map((p) => p.getBoundingClientRect().top * .75));
      expect(positions).toHaveLength(5);
      const expected = row.origins[kind];
      expect(Math.max(...positions.map((value, index) =>
        Math.abs((value - positions[0]) - (expected[index] - expected[0]))))).toBeLessThanOrEqual(.15);
      checks.push({ kind, positions });
    }
  };
  const download = async (name: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await pending).saveAs(root + '/' + name); outputs[name] = hash(await fs.readFile(root + '/' + name));
  };
  await check(); await download('source.pdf');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: new RegExp(`^Section 1 \\u2014 Default ${row.kind}`) }).click();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.insertText('Edited'); await expect(editor.locator('p').first()).toHaveText('Edited\tB');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText('A\tB');
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText('Edited\tB');
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText('A\tB');
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click(); await check();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `spacing-${row.name}`, exact: true }).click(); await check();
  await download('reloaded.pdf'); await download('original.docx', 'DOCX file Editable in Microsoft Word');
  expect(await fs.readFile(root + '/original.docx')).toEqual(source); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: row.docxHash, outputs, checks, errors,
    buildHash: await wordStoryBuildHash(), referenceHash: hash(await fs.readFile('tests/fixtures/native-word-story-spacing.json')),
    testHash: hash(await fs.readFile('tests/word-story-spacing.spec.ts')) }, null, 2));
});
