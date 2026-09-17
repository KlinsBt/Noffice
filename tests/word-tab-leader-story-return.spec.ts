import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-tab-leader-stories.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`native ${row.name} footer leader edit returns through history and reload`, async ({ page }) => {
  test.skip(process.env.NOFFICE_LEADER_STORY_RETURN !== '1', 'Requires independently edited installed Word output');
  const sourceRun = process.env.NOFFICE_LEADER_STORY_RETURN_SOURCE_RUN || 'exports-v1';
  if (!/^exports-v\d+$/.test(sourceRun)) throw Error('Invalid native source evidence folder');
  const root = `.local/word-tab-leader-stories/sources-v2/${row.name}/${sourceRun}`;
  const run = process.env.NOFFICE_LEADER_STORY_RETURN_RUN;
  if (run && !/^browser-v\d+$/.test(run)) throw Error('Invalid return evidence folder');
  const output = run ? root + '/' + run : root;
  await fs.mkdir(output, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile(root + '/returned.docx');
  const expected = JSON.parse(await fs.readFile(root + '/returned-glyphs.json', 'utf8'));
  expect(hash(source)).toBe(expected.docxHash);
  expect(hash(await fs.readFile(root + '/returned.pdf'))).toBe(expected.pdfHash);
  expect(hash(await fs.readFile(root + '/native-report.json'))).toBe(expected.nativeReceiptHash);
  const errors: string[] = [];page.on('pageerror', (e) => errors.push(e.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles({ name: `leader-return-${row.name}.docx`,
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const check = async () => {
    await expect(page.locator('.word-page-story [data-word-tab-leader-painted=true]')).toHaveCount(20);
    const footer = page.locator('.word-page-story[data-word-story*="footer"] p').first();
    await expect(footer).toHaveText('Native\tB');
    const actual = await footer.evaluate((p) => {
      const semantic = p.querySelector<HTMLElement>('[data-word-tab]')!;
      const tab = semantic.closest<HTMLElement>('[data-word-tab-measured]') || semantic;
      const scale = parseFloat(getComputedStyle(p.closest('.paper-wrap')!).zoom) || 1;
      const x = (tab.getBoundingClientRect().left - p.closest('.section-page')!.getBoundingClientRect().left) * .75 / scale;
      return (JSON.parse(tab.dataset.wordTabLeaderOffsets!) as number[]).map((offset) => x + offset * .75);
    });
    expect(actual).toHaveLength(expected.leaders.length);
    expect(Math.max(0, ...actual.map((x, i) => Math.abs(x - expected.leaders[i].x)))).toBeLessThanOrEqual(.15);
  };
  await check();await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: /^Section 1 \u2014 Default footer/ }).click();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await editor.focus();await page.keyboard.press('Control+Home');await page.keyboard.insertText('X');
  await expect(editor.locator('p').first()).toHaveText('XNative\tB');
  await page.keyboard.press('Control+z');await expect(editor.locator('p').first()).toHaveText('Native\tB');
  await page.keyboard.press('Control+y');await expect(editor.locator('p').first()).toHaveText('XNative\tB');
  await page.keyboard.press('Control+z');await expect(editor.locator('p').first()).toHaveText('Native\tB');
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();await check();
  await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `leader-return-${row.name}`, exact: true }).click();await check();
  await page.getByRole('button', { name: 'Export', exact: true }).click();const pending = page.waitForEvent('download');
  await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
  await (await pending).saveAs(output + '/browser-return.docx');
  expect(await fs.readFile(output + '/browser-return.docx')).toEqual(source);expect(errors).toEqual([]);
  await page.getByRole('button', { name: 'Export', exact: true }).click();const pdf = page.waitForEvent('download');
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();await (await pdf).saveAs(output + '/browser-return.pdf');
  expect(errors).toEqual([]);
  await fs.writeFile(output + '/browser-return-report.json', JSON.stringify({ sourceHash: hash(source), nativeReceiptHash: expected.nativeReceiptHash,
    pdfHash: hash(await fs.readFile(output + '/browser-return.pdf')),
    glyphReceiptHash: hash(await fs.readFile(root + '/returned-glyphs.json')), errors, buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-tab-leader-story-return.spec.ts')) }, null, 2));
});
