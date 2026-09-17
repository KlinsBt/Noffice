import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/word-pdf-decorations/returns.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout } from './word-pagination-helpers';

const run = process.env.NOFFICE_PDF_DECORATION_RETURNS_RUN || 'returns-browser-v1';
if (!/^returns-browser-v\d+$/.test(run)) throw Error('Invalid decoration return run');
const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');

for (const sample of reference.rows) test(`Word native ${sample.name} return, another formatting edit and files`, async ({ page }) => {
  test.setTimeout(90000);
  const root = `.local/word-pdf-decorations/${run}/${sample.name}`; await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile(sample.fixture); expect(hash(source)).toBe(sample.sourceHash);
  const outputs: Record<string, string> = {}, errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(sample.fixture);
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  await waitWordLayout(body); await expect(body).toHaveText(sample.snapshot.text.replace(/[\r\f]/g, ''));
  const capture = async (stage: string) => {
    for (const [extension, label] of [['docx', 'DOCX file Editable in Microsoft Word'], ['pdf', 'PDF file']]) {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download'); await page.getByRole('button', { name: label, exact: true }).click();
      const path = `${root}/${stage}.${extension}`; await (await pending).saveAs(path);
      outputs[`${stage}.${extension}`] = hash(await fs.readFile(path));
    }
  };
  await capture('source'); expect(outputs['source.docx']).toBe(sample.sourceHash);
  let editor = body;
  if (sample.name.startsWith('stories-')) {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: /^Section 1 .*Default header/ }).click();
    editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  }
  const bits = async () => editor.locator('p').first().evaluate(p => {
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), result: number[] = [];
    while (walker.nextNode()) { const text = walker.currentNode as Text;
      result.push(...Array.from(text.data, () => (text.parentElement?.closest('u') ? 1 : 0)
        | (text.parentElement?.closest('s,strike,del') ? 2 : 0))); }
    return result.slice(4, 23);
  });
  const native = sample.snapshot.decorations.find(row => !('empty' in row)
    && row.story === (sample.name.startsWith('stories-') ? 'Headers' : 'body'))!;
  const original = ('underline' in native && native.underline === 1 ? 1 : 0)
    | ('strike' in native && native.strike === -1 ? 2 : 0);
  await expect.poll(bits).toEqual(Array(19).fill(original));
  await editor.focus(); await page.keyboard.press('Control+Home');
  for (let i = 0; i < 4; i++) await page.keyboard.press('ArrowRight');
  for (let i = 0; i < 19; i++) await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Control+u'); await expect.poll(bits).toEqual(Array(19).fill(original ^ 1));
  await page.keyboard.press('Control+z'); await expect.poll(bits).toEqual(Array(19).fill(original));
  await page.keyboard.press('Control+y'); await expect.poll(bits).toEqual(Array(19).fill(original ^ 1));
  if (sample.name.startsWith('stories-')) await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await waitWordLayout(body); await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await capture('edited');
  const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click(); await waitWordLayout(body); await capture('reloaded');
  expect(errors).toEqual([]);
  await fs.writeFile(`${root}/report.json`, JSON.stringify({ sourceHash: sample.sourceHash, outputs, errors,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-pdf-decoration-returns.spec.ts')),
    scope: 'Actual native re-edit return import, original recovery, another browser underline edit/history/reload and downloads. Full native and direct PDF comparisons remain separate.' }, null, 2));
});
