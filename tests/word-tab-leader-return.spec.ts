import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

test('native leader edit returns through typing, history, reload and original export', async ({ page }) => {
  test.skip(process.env.NOFFICE_LEADER_NATIVE_RETURN !== '1', 'Requires independent installed Word outputs');
  const run = process.env.NOFFICE_LEADER_NATIVE_RUN || 'native-v1';
  const root = '.local/word-tab-leader-render', folder = root + '/' + run;
  const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
  const source = await fs.readFile(folder + '/returned.docx');
  const reference = JSON.parse(await fs.readFile(folder + '/returned-glyphs.json', 'utf8'));
  expect(reference.docxHash).toBe(hash(source));
  expect(reference.pdfHash).toBe(hash(await fs.readFile(folder + '/returned.pdf')));
  expect(reference.nativeReceiptHash).toBe(hash(await fs.readFile(folder + '/native-report.json')));
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: 'leader-native.docx', buffer: source,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  const first = body.locator('p').first();
  const match = async () => {
    await expect(first).toHaveText('Native\tB');
    await expect(body.locator('[data-word-tab-leader-painted=true]')).toHaveCount(90);
    await expect.poll(async () => {
      const actual = await first.locator('[data-word-tab]').evaluate((semanticTab) => {
        const tab = semanticTab.closest('[data-word-tab-measured]') || semanticTab;
        const host = tab.closest('[aria-label="Document text"]')!;
        const scale = parseFloat(getComputedStyle(tab.closest('.paper-wrap')!).zoom) || 1;
        return { x: (tab.getBoundingClientRect().left - host.getBoundingClientRect().left) * .75 / scale,
          offsets: JSON.parse(tab.getAttribute('data-word-tab-leader-offsets') || '[]') as number[] };
      });
      if (actual.offsets.length !== reference.leaders.length) return 10000;
      return Math.max(0, ...actual.offsets.map((x, i) => Math.abs(actual.x + x * .75 - reference.leaders[i].x)));
    }).toBeLessThanOrEqual(.15);
  };
  await match();await body.focus();await page.keyboard.press('Control+Home');await page.keyboard.insertText('X');
  await expect(first).toHaveText('XNative\tB');await page.keyboard.press('Control+z');await match();
  await page.keyboard.press('Control+y');await expect(first).toHaveText('XNative\tB');await page.keyboard.press('Control+z');await match();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'leader-native', exact: true }).click();await match();
  await page.getByRole('button', { name: 'Export', exact: true }).click();const pending = page.waitForEvent('download');
  await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
  const output = root + '/native-return.docx';await (await pending).saveAs(output);
  expect(await fs.readFile(output)).toEqual(source);
  await fs.writeFile(root + '/native-return-report.json', JSON.stringify({ sourceHash: hash(source),
    nativeReceiptHash: reference.nativeReceiptHash, glyphReceiptHash: hash(await fs.readFile(folder + '/returned-glyphs.json')),
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-tab-leader-return.spec.ts')) }, null, 2));
});
