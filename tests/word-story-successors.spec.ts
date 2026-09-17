import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout } from './word-pagination-helpers';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-story-successors.json', 'utf8'),
) as typeof import('./fixtures/native-word-story-successors.json');
const leading = JSON.parse(
  readFileSync('tests/fixtures/native-word-story-leading-deletion.json', 'utf8'),
) as typeof import('./fixtures/native-word-story-leading-deletion.json');
const cases = { ...reference.cases, ...leading.cases };
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

for (const [name, sample] of Object.entries(cases))
  test(`Word successor ${name} keeps native stories through body history and files`, async ({
    page,
  }) => {
    test.setTimeout(120000);
    await page.context().grantPermissions(['local-fonts']);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const root = `.local/word-story-successors/browser/${name}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(`tests/fixtures/${reference.sourceFile}`);
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
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `Successor ${name}.docx`,
      buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(page.locator('.section-page')).toHaveCount(9);
    let actions = 0;
    const link = async (section: number, kind = 'header', linked = false, slot = 'default') => {
      await openLinks(page, kind, slot);
      await page
        .getByLabel('Link section', { exact: true })
        .selectOption({ label: `Section ${section}` });
      await page
        .getByRole('checkbox', { name: 'Link to previous', exact: true })
        .setChecked(linked);
      await page.getByRole('button', { name: 'Apply header/footer link', exact: true }).click();
      actions++;
    };
    const edit = async (section: number, kind: string, text: string) => {
      const story = await openStory(page, section, kind, 'default');
      await story.focus();
      await page.keyboard.press('Control+Home');
      await page.keyboard.press('Shift+End');
      await page.keyboard.insertText(text);
      await expect(story).toHaveText(text);
      await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
      actions++;
    };
    let deletion:
      { middle?: boolean; two?: boolean; boundary?: boolean; partial?: boolean } | undefined;
    if (name in leading.cases) {
      if (name.includes('copy') || name === 'header-edit-first') {
        await link(
          2,
          name.startsWith('footer') ? 'footer' : 'header',
          false,
          name.startsWith('first-header')
            ? 'first'
            : name.startsWith('even-header')
              ? 'even'
              : 'default',
        );
        if (name === 'header-edit-first') await edit(2, 'header', 'Updated header');
      }
      deletion = {
        middle: name.includes('middle'),
        two: name.includes('two'),
        boundary: name.includes('boundary'),
        partial: name.includes('partial'),
      };
    } else if (name !== 'baseline') {
      const kind = name === 'unlink-middle-footer-edit' ? 'footer' : 'header';
      await link(2, kind);
      if (name !== 'unlink-middle-header') await edit(2, kind, `Updated ${kind}`);
      if (name === 'unlink-last-from-copy') {
        await link(3);
        await edit(3, 'header', 'Last header');
      }
      if (name === 'relink-middle') await link(2, 'header', true);
      if (name.startsWith('delete-')) deletion = { middle: name === 'delete-middle-after-copy' };
    }
    if (deletion) {
      await body.focus();
      await waitWordLayout(body);
      await body.evaluate((root, action) => {
        const paragraphs = [...root.querySelectorAll(':scope > p')];
        const end = action.middle || action.two ? 2 : 1;
        const range = document.createRange();
        if (action.boundary)
          range.setStart(paragraphs[end - 1], paragraphs[end - 1].childNodes.length);
        else if (action.partial) {
          const br = paragraphs[0].querySelector('br')!;
          range.setStartAfter(br);
        } else range.setStart(paragraphs[action.middle ? 1 : 0], 0);
        if (action.boundary) range.setEnd(paragraphs[end], 0);
        else range.setEnd(paragraphs[end], 0);
        const selection = window.getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
        document.dispatchEvent(new Event('selectionchange'));
      }, deletion);
      await page.keyboard.press('Backspace');
      actions++;
    }
    const layout = async () => {
      await expect(page.locator('.section-page')).toHaveCount(sample.native.pages);
      await waitWordLayout(body);
      await expect(body).toHaveText(sample.native.body.replace(/[\r\n\v\f]/g, ''));
      for (const [index, native] of sample.paintPages.entries()) {
        const text = native.lines
          .filter((l) => l.text.includes('header') || l.text.includes('footer'))
          .map((l) => l.text);
        await expect(
          page.locator('.section-page').nth(index).locator('.word-page-story'),
        ).toHaveText([
          text.filter((t) => t.includes('header')).join(''),
          text.filter((t) => t.includes('footer')).join(''),
        ]);
      }
    };
    await layout();
    await capture('edited', true);
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    if (actions) {
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      if (deletion) await expect(page.locator('.section-page')).toHaveCount(9);
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await layout();
    }
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: `Successor ${name}`, exact: true }).click();
    await layout();
    await capture('reloaded');
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/edited.docx`);
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'edited',
    );
    await layout();
    await capture('reimported');
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify({ hashes, errors, buildHash: await wordStoryBuildHash() }, null, 2),
    );
  });
