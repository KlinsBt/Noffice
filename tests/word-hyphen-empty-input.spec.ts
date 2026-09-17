import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import cases from './fixtures/word-hyphen-empty-input/cases.json' with { type: 'json' };
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
for (const sample of cases.rows) test(`Empty Word paragraph ${sample.name}: font, input, history and actual files`, async ({ page }) => {
  const run = process.env.NOFFICE_EMPTY_HYPHEN_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid empty-input evidence run');
  const root = `.local/word-hyphen-empty-input/${run}/${sample.name}`; await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile(sample.path); expect(hash(source)).toBe(sample.sourceHash);
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: sample.name + '.docx', buffer: source,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true }); await waitWordLayout(editor);
  const model = () => editor.evaluate(el => (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getJSON());
  const verify = async () => {
    const p = await editor.evaluate(el => (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.getJSON().content![0]);
    expect(p.content).toHaveLength(1); const node = p.content![0];
    expect(node.type).toBe('wordHyphen');
    expect('attrs' in node ? node.attrs?.kind : undefined).toBe(sample.kind);
    expect(node.marks?.find(mark => mark.type === 'textStyle')?.attrs)
      .toMatchObject({ fontFamily: sample.family, fontSize: sample.size + 'pt' });
    await expect(editor.locator('p').nth(1)).toHaveText('Following paragraph.');
    const painted = await editor.locator('[data-word-hyphen]').evaluate(el => ({ family: getComputedStyle(el).fontFamily,
      size: parseFloat(getComputedStyle(el).fontSize) }));
    expect(painted.family.replaceAll('"', '')).toBe(sample.family); expect(painted.size).toBeCloseTo(sample.size * 4 / 3, 3);
  };
  await editor.evaluate(el => (el as HTMLElement & { editor: import('@tiptap/core').Editor }).editor.commands.setTextSelection(1));
  const before = await model(); await editor.focus();
  if (sample.kind === 'optional') await page.keyboard.press('Control+-');
  else await page.keyboard.insertText('\u00ad');
  await waitWordLayout(editor); await verify(); const inserted = await model();
  await page.keyboard.press('Control+z'); await expect.poll(model).toEqual(before);
  await page.keyboard.press('Control+y'); await expect.poll(model).toEqual(inserted); await waitWordLayout(editor);
  const hashes: Record<string, string> = {};
  const capture = async (stage: string) => {
    for (const [extension, label] of [['docx', 'DOCX file Editable in Microsoft Word'], ['pdf', 'PDF file']]) {
      await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: label, exact: true }).click(); const path = `${root}/${stage}.${extension}`;
      await (await pending).saveAs(path); hashes[stage + '.' + extension] = hash(await fs.readFile(path));
    }
  };
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible(); await capture('inserted');
  const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click(); await waitWordLayout(editor); await verify(); await capture('reloaded');
  await page.locator('input[type=file][multiple]').setInputFiles(root + '/inserted.docx'); await waitWordLayout(editor); await verify();
  await capture('reimported'); expect(hashes['reimported.docx']).toBe(hashes['inserted.docx']);
  await fs.writeFile(root + '/report.json', JSON.stringify({ name: sample.name, sourceHash: sample.sourceHash,
    before, inserted, hashes, buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-hyphen-empty-input.spec.ts')),
    scope: 'Actual optional shortcut/literal text input into an empty exact24 Arial10/Calibri11 paragraph; inherited font, full-model history, reload, DOCX/PDF downloads and reimport. Native rendering and physical-keyboard acceptance remain separate.' }, null, 2));
});
