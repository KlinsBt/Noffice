import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout, wordLineOrigins, wordPlacements } from './word-pagination-helpers';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-section-stories.json', 'utf8'),
) as typeof import('./fixtures/native-word-section-stories.json');
test.describe.configure({ mode: 'parallel' });

for (const [name, sample] of Object.entries(reference.cases))
  test(`Word ${name} renders section stories and preserves their edits across real files`, async ({
    page,
  }) => {
    test.setTimeout(90000);
    await page.context().grantPermissions(['local-fonts']);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const root = `.local/word-section-stories/browser/${name}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(`tests/fixtures/${sample.sourceFile}`);
    const hashes: Record<string, string> = {
      source: createHash('sha256').update(source).digest('hex'),
    };
    const capture = async (stage: string, pdf = false) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const waiting = page.waitForEvent('download');
      await page
        .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
        .click();
      await (await waiting).saveAs(`${root}/${stage}.docx`);
      hashes[`${stage}.docx`] = createHash('sha256')
        .update(await fs.readFile(`${root}/${stage}.docx`))
        .digest('hex');
      if (pdf) await capturePdf(page, root, stage, hashes);
    };
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `${name}.docx`,
      buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(page.locator('.section-page')).toHaveCount(sample.native.pages);
    await waitWordLayout(body);
    const originalLines = (await wordLineOrigins(body)).flat();
    expect(originalLines.map((l) => l.page)).toEqual(
      sample.paintPages.flatMap((p) =>
        p.lines.filter((l) => l.text.startsWith('B')).map(() => p.page),
      ),
    );
    const placements = (await wordPlacements(body)).flat();
    for (const p of sample.paintPages)
      expect(
        placements
          .filter((c) => c.page === p.page)
          .map((c) => c.text)
          .join('')
          .replace(/\s/g, ''),
      ).toBe(
        p.lines
          .filter((l) => l.text.startsWith('B'))
          .map((l) => l.text)
          .join('')
          .replace(/\s/g, ''),
      );
    for (const [index, native] of sample.paintPages.entries()) {
      const stories = page.locator('.section-page').nth(index).locator('.word-page-story');
      if (!native.lines.length) await expect(stories).toHaveCount(0);
      else
        await expect(stories).toHaveText([
          native.lines
            .filter((l) => !l.text.startsWith('B') && !l.text.includes('footer'))
            .map((l) => l.text)
            .join(''),
          native.lines
            .filter((l) => l.text.includes('footer'))
            .map((l) => l.text)
            .join(''),
        ]);
    }
    await capturePdf(page, root, 'source', hashes);
    await page.pdf({ path: `${root}/print.pdf`, preferCSSPageSize: true });
    hashes['print.pdf'] = createHash('sha256')
      .update(await fs.readFile(`${root}/print.pdf`))
      .digest('hex');
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: /^Section 2 — Default header/ }).click();
    const story = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    await story.focus();
    await page.keyboard.press('Control+Home');
    await page.keyboard.press('Control+Shift+End');
    await page.keyboard.insertText('Changed second header');
    await expect(story).toHaveText('Changed second header');
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    const changed = page.locator('.word-page-story').filter({ hasText: 'Changed second header' });
    await expect(changed.first()).toBeVisible();
    await capture('edited', true);
    await waitWordLayout(body);
    const editedLines = (await wordLineOrigins(body)).flat();
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(changed).toHaveCount(0);
    await capture('undo');
    expect(hashes['undo.docx']).toBe(hashes.source);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(changed.first()).toBeVisible();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    await expect(changed.first()).toBeVisible();
    await capture('reloaded');
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/reloaded.docx`);
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'reloaded',
    );
    await expect(changed.first()).toBeVisible();
    await waitWordLayout(body);
    expect((await wordLineOrigins(body)).flat()).toEqual(editedLines);
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify(
        { hashes, errors, originalLines, editedLines, buildHash: await wordStoryBuildHash() },
        null,
        2,
      ),
    );
  });

async function capturePdf(page: Page, root: string, stage: string, hashes: Record<string, string>) {
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const waiting = page.waitForEvent('download');
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await (await waiting).saveAs(`${root}/${stage}.pdf`);
  hashes[`${stage}.pdf`] = createHash('sha256')
    .update(await fs.readFile(`${root}/${stage}.pdf`))
    .digest('hex');
}
