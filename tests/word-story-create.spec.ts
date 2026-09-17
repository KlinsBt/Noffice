import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout } from './word-pagination-helpers';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-story-create.json', 'utf8'),
) as typeof import('./fixtures/native-word-story-create.json');
test.describe.configure({ mode: 'parallel' });
async function openLinks(page: Page, kind: string, slot: string) {
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: 'Section links', exact: true }).click();
  await page.getByLabel('Header or footer', { exact: true }).selectOption(kind);
  await page.getByLabel('Header/footer page type', { exact: true }).selectOption(slot);
}
async function openStory(page: Page, section: number, kind: string, slot: string) {
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  const label = slot === 'first' ? 'First-page' : slot === 'even' ? 'Even-page' : 'Default';
  await page
    .getByRole('button', { name: new RegExp(`^Section ${section} \u2014 ${label} ${kind}`) })
    .click();
  return page.getByRole('textbox', { name: 'Header or footer text', exact: true });
}

for (const [name, sample] of Object.entries(reference.cases))
  test(`Word creates ${name} through history, reload and actual files`, async ({ page }) => {
    test.setTimeout(120000);
    await page.context().grantPermissions(['local-fonts']);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const root = `.local/word-story-create/browser/${name}`;
    await fs.mkdir(root, { recursive: true });
    const sourceInfo = reference.sources[sample.mode as keyof typeof reference.sources];
    const source = await fs.readFile('tests/fixtures/' + sourceInfo.file);
    const hashes: Record<string, string> = {
      source: createHash('sha256').update(source).digest('hex'),
    };
    const capture = async (stage: string, pdf = false) => {
      for (const ext of pdf ? ['docx', 'pdf'] : ['docx']) {
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        const waiting = page.waitForEvent('download');
        await page
          .getByRole('button', {
            name: ext === 'docx' ? 'DOCX file Editable in Microsoft Word' : 'PDF file',
            exact: true,
          })
          .click();
        const path = `${root}/${stage}.${ext}`;
        await (await waiting).saveAs(path);
        hashes[`${stage}.${ext}`] = createHash('sha256')
          .update(await fs.readFile(path))
          .digest('hex');
      }
    };
    await page.goto('/');
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({
        name: `Create ${name}.docx`,
        buffer: source,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(page.locator('.section-page')).toHaveCount(sample.blank.pages);
    const kind = sample.kind === 'Headers' ? 'header' : 'footer',
      slot = ({ 1: 'default', 2: 'first', 3: 'even' } as const)[sample.slot as 1 | 2 | 3];
    let story = await openStory(page, sample.section, kind, slot);
    await expect(story).toHaveText('');
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await capture('untouched');
    expect(hashes['untouched.docx']).toBe(hashes.source);
    if (sample.section === 2 && !('linked' in sample && sample.linked)) {
      await openLinks(page, kind, slot);
      await page.getByLabel('Link section', { exact: true }).selectOption({ label: 'Section 2' });
      await page.getByRole('checkbox', { name: 'Link to previous', exact: true }).uncheck();
      await page.getByRole('button', { name: 'Apply header/footer link', exact: true }).click();
      await page.getByRole('button', { name: 'Home', exact: true }).click();
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await capture('unlink-undone');
      expect(hashes['unlink-undone.docx']).toBe(hashes.source);
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
    }
    await capture('blank');
    story = await openStory(page, sample.section, kind, slot);
    await story.focus();
    await page.keyboard.press('Control+Home');
    await page.keyboard.insertText(sample.text);
    await expect(story).toHaveText(sample.text);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    story = await openStory(page, sample.section, kind, slot);
    await expect(story).toHaveText('');
    await story.focus();
    await page.keyboard.press('Control+Home');
    await page.keyboard.insertText(sample.text);
    const toolbar = page.getByRole('toolbar', { name: 'Header or footer formatting', exact: true });
    await toolbar.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(story).toHaveText('');
    await toolbar.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(story).toHaveText(sample.text);
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    const layout = async (pages: typeof sample.paintPages.typed) => {
      await expect(page.locator('.section-page')).toHaveCount(pages.length);
      await waitWordLayout(body);
      await expect(body).toHaveText(sample.typed.body.replace(/[\r\n\v\f]/g, ''));
      for (const [index, expected] of pages.entries())
        await expect(
          page.locator('.section-page').nth(index).locator('.word-page-story'),
        ).toHaveText([
          expected.lines
            .filter((l) => l.text.startsWith('Created headers'))
            .map((l) => l.text)
            .join(''),
          expected.lines
            .filter((l) => l.text.startsWith('Created footers'))
            .map((l) => l.text)
            .join(''),
        ]);
    };
    await layout(sample.paintPages.typed);
    await capture('typed', true);
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(await openStory(page, sample.section, kind, slot)).toHaveText('');
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: `Create ${name}`, exact: true }).click();
    await layout(sample.paintPages.typed);
    await capture('reloaded');
    await expect(await openStory(page, sample.section, kind, slot)).toHaveText(sample.text);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/typed.docx`);
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'typed',
    );
    await layout(sample.paintPages.typed);
    await capture('reimported');
    expect(hashes['reimported.docx']).toBe(hashes['typed.docx']);
    await expect(await openStory(page, sample.section, kind, slot)).toHaveText(sample.text);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    if ('activation' in sample) {
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      await page.getByRole('button', { name: 'Page options', exact: true }).click();
      await page
        .getByLabel('Header/footer section', { exact: true })
        .selectOption({ label: `Section ${sample.section}` });
      await page
        .getByRole('checkbox', {
          name: slot === 'first' ? 'Different first page' : 'Different odd and even pages',
          exact: true,
        })
        .check();
      await page.getByRole('button', { name: 'Apply page options', exact: true }).click();
      await layout(sample.activation.paintPages);
      await capture('activated', true);
      await page.getByRole('button', { name: 'Home', exact: true }).click();
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await layout(sample.paintPages.typed);
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await layout(sample.activation.paintPages);
    }
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify({ hashes, errors, buildHash: await wordStoryBuildHash() }, null, 2),
    );
  });
