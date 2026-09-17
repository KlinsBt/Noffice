import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-story-links.json', 'utf8'),
) as typeof import('./fixtures/native-word-story-links.json');
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout, wordLineOrigins } from './word-pagination-helpers';

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
  test(`Word ${name} preserves shared or independent editing through history and actual files`, async ({
    page,
  }) => {
    test.setTimeout(90000);
    await page.context().grantPermissions(['local-fonts']);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const root = `.local/word-story-links/browser/${name}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(`tests/fixtures/${sample.sourceFile}`);
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
    const { kind, slot, linked } = sample.action;
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `Links ${name}.docx`,
      buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(page.locator('.section-page')).toHaveCount(sample.native.pages);
    const bodyText = await body.innerText();
    await openLinks(page, kind, slot);
    await page.getByLabel('Link section', { exact: true }).selectOption({ label: 'Section 1' });
    await expect(
      page.getByRole('checkbox', { name: 'Link to previous', exact: true }),
    ).toBeDisabled();
    await expect(
      page.getByRole('button', { name: 'Apply header/footer link', exact: true }),
    ).toBeDisabled();
    await page.getByLabel('Link section', { exact: true }).selectOption({ label: 'Section 2' });
    const toggle = page.getByRole('checkbox', { name: 'Link to previous', exact: true });
    await expect(toggle).toBeChecked({ checked: !linked });
    await toggle.setChecked(linked);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.getByRole('button', { name: 'Close dialog', exact: true }).click();
    await openLinks(page, kind, slot);
    await expect(toggle).toBeChecked({ checked: !linked });
    await toggle.setChecked(linked);
    await page.getByRole('button', { name: 'Apply header/footer link', exact: true }).click();
    const layout = async () => {
      await expect(page.locator('.section-page')).toHaveCount(sample.native.pages);
      await waitWordLayout(body);
      await expect
        .poll(async () => (await wordLineOrigins(body)).flat().map((line) => line.page))
        .toEqual(
          sample.paintPages.flatMap((p) =>
            p.lines.filter((line) => line.text.startsWith('B')).map(() => p.page),
          ),
        );
      await expect(body).toHaveText(bodyText, { useInnerText: true });
    };
    await layout();
    for (const [index, native] of sample.paintPages.entries())
      await expect(page.locator('.section-page').nth(index).locator('.word-page-story')).toHaveText(
        [
          native.lines
            .filter((l) => l.text.includes('header'))
            .map((l) => l.text)
            .join(''),
          native.lines
            .filter((l) => l.text.includes('footer'))
            .map((l) => l.text)
            .join(''),
        ],
      );
    await capture('linked', true);
    await openLinks(page, kind, slot);
    await expect(toggle).toBeChecked({ checked: linked });
    await page.getByRole('button', { name: 'Apply header/footer link', exact: true }).click();
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await capture('undo');
    expect(hashes['undo.docx']).toBe(hashes.source);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: `Links ${name}`, exact: true }).click();
    await layout();
    await capture('reloaded');
    const story = await openStory(page, 2, kind, slot);
    const previous = await story.innerText(),
      edited = `Updated ${kind}`;
    await story.focus();
    await page.keyboard.press('Control+Home');
    await page.keyboard.press('Shift+End');
    await page.keyboard.insertText(edited);
    await expect(story).toHaveText(edited);
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await layout();
    await capture('edited', true);
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(await openStory(page, 2, kind, slot)).toHaveText(previous);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(await openStory(page, 1, kind, slot)).toHaveText(linked ? edited : previous);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/edited.docx`);
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'edited',
    );
    await layout();
    await expect(await openStory(page, 2, kind, slot)).toHaveText(edited);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await expect(await openStory(page, 1, kind, slot)).toHaveText(linked ? edited : previous);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify({ hashes, errors, buildHash: await wordStoryBuildHash() }, null, 2),
    );
  });
