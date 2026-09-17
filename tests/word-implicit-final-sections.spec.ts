import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import reference from './fixtures/native-word-implicit-final-sections.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout } from './word-pagination-helpers';

const hash = (value: Buffer) => createHash('sha256').update(value).digest('hex');
const run = process.env.NOFFICE_IMPLICIT_SECTIONS_RUN || 'browser-v1';
if (!/^browser-v\d+$/.test(run)) throw Error('Invalid implicit-section run');

for (const sample of reference.rows) test(`Word ${sample.name} import, editing, history and retained files`, async ({ page }) => {
  test.setTimeout(120000);
  const root = `.local/word-implicit-final-sections/${run}/${sample.name}`;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile(sample.fixture);
  expect(hash(source)).toBe(sample.sourceHash);
  const outputs: Record<string, string> = {}, errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);
  await page.setViewportSize({ width: 1600, height: 1200 });
  await page.goto('/');
  const name = 'Implicit final ' + sample.name;
  const upload = async (fileName: string, buffer: Buffer) => {
    await page.locator('input[type=file][multiple]').setInputFiles({ name: fileName + '.docx', buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(fileName);
  };
  await upload(name, source);
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  const native = sample.states.find(state => state.stage === 'source')!.sourceSnapshot;
  // A short, single section retains its exact source paper; multi-section
  // documents use the measured per-page surfaces. Assert both actual papers.
  const papers = page.locator('.paper-wrap.source-layout > .paper, .section-page');
  const layout = async (edited: boolean) => {
    await waitWordLayout(body);
    await expect(papers).toHaveCount(native.pages);
    await expect(body).toHaveText((edited ? 'Edited ' : '') + native.text.replace(/[\r\f]/g, ''));
    const geometry = await papers.evaluateAll(elements => elements.map(element => {
      const box = element.getBoundingClientRect();
      return { width: box.width * .75, height: box.height * .75,
        stories: [...element.querySelectorAll('.word-page-story')].map(story => story.textContent) };
    }));
    expect(geometry).toHaveLength(native.sections.length);
    geometry.forEach((page, index) => {
      expect(page.width).toBeCloseTo(native.sections[index].width, 1);
      expect(page.height).toBeCloseTo(native.sections[index].height, 1);
    });
    return geometry;
  };
  const capture = async (stage: string) => {
    for (const [extension, label] of [['docx', 'DOCX file Editable in Microsoft Word'], ['pdf', 'PDF file']]) {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: label, exact: true }).click();
      const file = `${root}/${stage}.${extension}`;
      await (await pending).saveAs(file);
      outputs[`${stage}.${extension}`] = hash(await fs.readFile(file));
    }
  };
  const initial = await layout(false);
  await capture('original');
  expect(outputs['original.docx']).toBe(sample.sourceHash);
  await body.focus(); await page.keyboard.press('Control+Home');
  await page.keyboard.type('Edited ');
  const edited = await layout(true);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await layout(false);
  await expect(page.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await layout(true);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await capture('edited');
  await page.pdf({ path: `${root}/edited-print.pdf`, preferCSSPageSize: true });
  outputs['edited-print.pdf'] = hash(await fs.readFile(`${root}/edited-print.pdf`));
  await layout(true);
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click();
  await layout(true); await capture('reloaded');
  const output = await fs.readFile(`${root}/edited.docx`);
  await upload('Implicit reimport ' + sample.name, output);
  await layout(true); await capture('reimported');
  const originalZip = await JSZip.loadAsync(source);
  for (const stage of ['edited', 'reloaded', 'reimported']) {
    const zip = await JSZip.loadAsync(await fs.readFile(`${root}/${stage}.docx`));
    for (const file of Object.keys(originalZip.files).filter(file => !originalZip.files[file].dir && file !== 'word/document.xml'))
      expect(await zip.file(file)?.async('nodebuffer'), file).toEqual(await originalZip.file(file)!.async('nodebuffer'));
  }
  expect(errors).toEqual([]);
  await fs.writeFile(`${root}/browser-report.json`, JSON.stringify({ name: sample.name, sourceHash: sample.sourceHash,
    outputs, initial, edited, errors, buildHash: await wordStoryBuildHash(),
    testHash: hash(await fs.readFile('tests/word-implicit-final-sections.spec.ts')),
    scope: 'Import/text insertion/history/reload/reimport/actual files only. Boundary keyboard behavior and failure matrix remain separate required cases.' }, null, 2));
});
