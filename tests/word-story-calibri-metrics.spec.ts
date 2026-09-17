import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout } from './word-pagination-helpers';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-story-calibri-metrics.json', 'utf8'),
) as typeof import('./fixtures/native-word-story-calibri-metrics.json');
test.describe.configure({ mode: 'parallel' });
test.skip(
  process.env.NOFFICE_STORY_CALIBRI_NATIVE !== '1',
  'Requires hash-bound local native Calibri discovery specimens.',
);
for (const sample of reference.rows)
  test(`Word retains native Calibri story ${sample.name} through editing, history and reload`, async ({
    page,
  }) => {
    await page.context().grantPermissions(['local-fonts']);
    const root = `.local/word-story-calibri-metrics/browser/${sample.name}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(`.local/word-story-calibri-metrics/${sample.name}.docx`);
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    expect(hash(source)).toBe(sample.docxHash);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const hashes: Record<string, string> = { source: hash(source) };
    await page.goto('/');
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({
        name: `Calibri ${sample.name}.docx`,
        buffer: source,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(page.locator('.section-page')).toHaveCount(sample.native.pages);
    await waitWordLayout(body);
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page
      .getByRole('button', {
        name: `Section 1 — Default ${sample.kind === 'Headers' ? 'header' : 'footer'}`,
        exact: true,
      })
      .click();
    const story = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    const before = await story.textContent();
    await story.focus();
    await page.keyboard.press('Control+End');
    await page.keyboard.insertText('!');
    await expect(story).toHaveText(before + '!');
    const toolbar = page.getByRole('toolbar', { name: 'Header or footer formatting', exact: true });
    await toolbar.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(story).toHaveText(before!);
    await toolbar.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(story).toHaveText(before + '!');
    await toolbar.getByRole('button', { name: 'Undo', exact: true }).click();
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    for (const stage of ['source', 'reloaded']) {
      if (stage === 'reloaded') {
        await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
        await page.reload();
        await page.getByRole('button', { name: 'Recent files', exact: true }).click();
        await page.getByRole('button', { name: `Calibri ${sample.name}`, exact: true }).click();
        await expect(page.locator('.section-page')).toHaveCount(sample.native.pages);
        await waitWordLayout(body);
      }
      for (const ext of ['docx', 'pdf']) {
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        const pending = page.waitForEvent('download');
        await page
          .getByRole('button', {
            name: ext === 'pdf' ? 'PDF file' : 'DOCX file Editable in Microsoft Word',
            exact: true,
          })
          .click();
        const path = `${root}/${stage}.${ext}`;
        await (await pending).saveAs(path);
        hashes[`${stage}.${ext}`] = hash(await fs.readFile(path));
        if (ext === 'docx') expect(hashes[`${stage}.${ext}`]).toBe(hashes.source);
      }
    }
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify({ hashes, errors, buildHash: await wordStoryBuildHash() }, null, 2),
    );
  });
