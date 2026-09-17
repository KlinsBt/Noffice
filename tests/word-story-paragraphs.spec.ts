import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout, wordLineOrigins } from './word-pagination-helpers';

async function open(page: Page, kind: string) {
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: `Section 1 — Default ${kind}`, exact: true }).click();
  return page.getByRole('textbox', { name: 'Header or footer text', exact: true });
}
for (const kind of ['header', 'footer'])
  test(`Word ${kind} paragraph creation, join and deletion survive history and retained export`, async ({
    page,
  }) => {
    test.setTimeout(90000);
    await page.context().grantPermissions(['local-fonts']);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const root = `.local/word-side-stories/paragraphs/${kind}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile('tests/fixtures/word-side-stories/default.docx');
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
      hashes[stage] = createHash('sha256')
        .update(await fs.readFile(`${root}/${stage}.docx`))
        .digest('hex');
      if (pdf) {
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        const waiting = page.waitForEvent('download');
        await page.getByRole('button', { name: 'PDF file', exact: true }).click();
        await (await waiting).saveAs(`${root}/${stage}.pdf`);
        hashes[`${stage}.pdf`] = createHash('sha256')
          .update(await fs.readFile(`${root}/${stage}.pdf`))
          .digest('hex');
      }
    };
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `Paragraph ${kind}.docx`,
      buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.locator('.section-page')).toHaveCount(4);
    const story = await open(page, kind);
    await story.focus();
    await page.keyboard.press('Control+End');
    for (let i = 1; i <= 3; i++) {
      await page.keyboard.press('Enter');
      await page.keyboard.insertText(`extra${i}`);
    }
    await expect(story.locator('p')).toHaveCount(4);
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await expect(page.locator('.section-page')).toHaveCount(5);
    await capture('edited', true);
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(page.locator('.section-page')).toHaveCount(4);
    await capture('undo');
    expect(hashes.undo).toBe(hashes.source);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(page.locator('.section-page')).toHaveCount(5);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: `Paragraph ${kind}`, exact: true }).click();
    await expect(page.locator('.section-page')).toHaveCount(5);
    await capture('reloaded');
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/reloaded.docx`);
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'reloaded',
    );
    await expect(page.locator('.section-page')).toHaveCount(5);
    await open(page, kind);
    await expect(story.locator('p')).toHaveCount(4);
    await story.focus();
    await page.keyboard.press('Control+Home');
    await page.keyboard.press('End');
    await page.keyboard.press('Delete');
    await expect(story.locator('p')).toHaveCount(3);
    await page.keyboard.press('Control+End');
    await page.keyboard.press('Shift+Home');
    await page.keyboard.press('Backspace');
    await page.keyboard.press('Backspace');
    await expect(story.locator('p')).toHaveText([
      `${kind === 'header' ? 'Header' : 'Footer'} defaultextra1`,
      'extra2',
    ]);
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await expect(page.locator('.section-page')).toHaveCount(4);
    await capture('joined', true);
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(page.locator('.section-page')).toHaveCount(5);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(page.locator('.section-page')).toHaveCount(4);
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await waitWordLayout(body);
    const lines = (await wordLineOrigins(body)).flat();
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/joined.docx`);
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'joined',
    );
    await expect(page.locator('.section-page')).toHaveCount(4);
    await waitWordLayout(body);
    expect((await wordLineOrigins(body)).flat().map((l) => l.page)).toEqual(
      lines.map((l) => l.page),
    );
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify({ hashes, errors, lines, buildHash: await wordStoryBuildHash() }, null, 2),
    );
  });
