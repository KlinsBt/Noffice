import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-story-spacing-grid.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`native story grid ${row.name}: further edit returns through history and reload`, async ({ page }) => {
  test.skip(process.env.NOFFICE_STORY_GRID_RETURN !== '1', 'Requires independently edited installed Word output');
  const run = process.env.NOFFICE_STORY_GRID_RETURN_RUN || 'browser-v1';
  const sourceRun = process.env.NOFFICE_STORY_GRID_SOURCE_RUN || 'grid-browser-v3';
  if (!/^browser-v\d+$/.test(run) || !/^grid-browser-v\d+$/.test(sourceRun)) throw Error('Invalid native return evidence folder');
  const root = `.local/word-story-spacing/${sourceRun}/${row.name}/native-v1`, output = root + '/' + run;
  await fs.mkdir(output, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile(root + '/returned.docx');
  const nativeBytes = await fs.readFile(root + '/native-report.json');
  const native = JSON.parse(nativeBytes.toString().replace(/^\uFEFF/, ''));
  const expected = native.stages.find((stage: { stage: string }) => stage.stage === 'returned');
  expect(hash(source)).toBe(expected.docxHash);
  expect(hash(await fs.readFile(root + '/returned.pdf'))).toBe(expected.pdfHash);
  expect(native.executableHash).toBe(reference.executableHash);
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: `story-grid-return-${row.name}.docx`,
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const story = page.locator(`.word-page-story[data-word-story*="${row.kind}"]`).first();
  const expectedText = expected.paragraphs[0].text.replace(/\r$/, '');
  const check = async () => {
    await expect(page.locator('.section-page')).toHaveCount(2);
    await expect(story.locator('p').first()).toHaveText(expectedText);
    await expect.poll(async () => {
      const values = await story.locator('p').evaluateAll(paragraphs => paragraphs.map(p => {
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), values = [];
      while (walker.nextNode()) {
        const node = walker.currentNode as Text;
        if (!node.data || node.parentElement!.closest('.ProseMirror-widget,[contenteditable=false]')) continue;
        const style = getComputedStyle(node.parentElement!);
        values.push({ family: style.fontFamily.replace(/^['"]|['"]$/g, ''), size: parseFloat(style.fontSize) * .75 });
      }
      return values;
      }));
      return values.length === expected.paragraphs.length && values.every((runs, index) =>
        runs.length > 0 && runs.every(run => run.family === expected.paragraphs[index].family &&
          Math.abs(run.size - expected.paragraphs[index].size) < .001));
    }).toBe(true);
  };
  await check(); await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: new RegExp(`^Section 1 \u2014 Default ${row.kind}`) }).click();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.insertText('X');
  await expect(editor.locator('p').first()).toHaveText('X' + expectedText);
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText(expectedText);
  await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText('X' + expectedText);
  await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText(expectedText);
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click(); await check();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `story-grid-return-${row.name}`, exact: true }).click(); await check();
  const outputs: Record<string, string> = {};
  for (const [filename, label] of [['returned.docx', 'DOCX file Editable in Microsoft Word'], ['returned.pdf', 'PDF file']]) {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();
    await (await pending).saveAs(output + '/' + filename);
    outputs[filename] = hash(await fs.readFile(output + '/' + filename));
  }
  expect(await fs.readFile(output + '/returned.docx')).toEqual(source); expect(errors).toEqual([]);
  await fs.writeFile(output + '/browser-report.json', JSON.stringify({ sourceHash: hash(source),
    nativeReceiptHash: hash(nativeBytes), outputs, errors, buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-story-spacing-grid-returns.spec.ts')) }, null, 2));
});
