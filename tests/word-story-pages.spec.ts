import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
import { readFileSync } from 'node:fs';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-side-stories.json', 'utf8'),
) as typeof import('./fixtures/native-word-side-stories.json');
import { wordLineOrigins, waitWordLayout } from './word-pagination-helpers';

test.describe.configure({ mode: 'parallel' });
for (const [name, sample] of Object.entries(reference.cases))
  test(`Word ${name} renders page-selected stories and exports their actual PDF`, async ({
    page,
  }) => {
    test.setTimeout(90000);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.context().grantPermissions(['local-fonts']);
    const root = `.local/word-side-stories/render/${name}`;
    await fs.mkdir(root, { recursive: true });
    await page.goto('/');
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles(`tests/fixtures/${sample.sourceFile}`);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(page.locator('.section-page')).toHaveCount(sample.native.pages);
    await waitWordLayout(editor);
    const lines = (await wordLineOrigins(editor)).flat();
    expect(lines.map((l) => l.page)).toEqual(
      sample.paintPages.flatMap((p) =>
        p.lines.filter((l) => l.text.startsWith('body')).map(() => p.page),
      ),
    );
    const stories = await page
      .locator('.section-page')
      .evaluateAll((pages) =>
        pages.map((p) =>
          [...p.querySelectorAll<HTMLElement>('.word-page-story')].map((s) => s.innerText),
        ),
      );
    expect(stories).toHaveLength(sample.native.pages);
    for (const [i, values] of stories.entries()) {
      expect(values.map((s) => s.replace(/\s+/g, ' ').trim())).toEqual([
        sample.paintPages[i].lines
          .filter((l) => /header/i.test(l.text))
          .map((l) => l.text)
          .join(' '),
        sample.paintPages[i].lines
          .filter((l) => /footer/i.test(l.text))
          .map((l) => l.text)
          .join(' '),
      ]);
    }
    await page.screenshot({ path: `${root}/screen.png`, fullPage: true });
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: 'PDF file', exact: true }).click();
    await (await waiting).saveAs(`${root}/download.pdf`);
    await page.pdf({ path: `${root}/print.pdf`, preferCSSPageSize: true });
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify(
        {
          lines,
          buildHash: await wordStoryBuildHash(),
          sourceSha256: createHash('sha256')
            .update(await fs.readFile(`tests/fixtures/${sample.sourceFile}`))
            .digest('hex'),
          stories,
          errors,
          hashes: Object.fromEntries(
            await Promise.all(
              ['download.pdf', 'print.pdf'].map(async (path) => [
                path,
                createHash('sha256')
                  .update(await fs.readFile(`${root}/${path}`))
                  .digest('hex'),
              ]),
            ),
          ),
        },
        null,
        2,
      ),
    );
  });
