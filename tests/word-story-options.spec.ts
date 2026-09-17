import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
import { wordLineOrigins, waitWordLayout } from './word-pagination-helpers';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-story-options.json', 'utf8'),
) as typeof import('./fixtures/native-word-story-options.json');
test.describe.configure({ mode: 'parallel' });
const openOptions = async (page: Page) => {
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: 'Page options', exact: true }).click();
  await expect(
    page.getByRole('dialog', { name: 'Header and footer page options', exact: true }),
  ).toBeVisible();
};
for (const [name, sample] of Object.entries(reference.cases))
  test(`Word ${name} page options survive history, reload and actual exported files`, async ({
    page,
  }) => {
    test.setTimeout(90000);
    await page.context().grantPermissions(['local-fonts']);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const root = `.local/word-story-options/browser/${name}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(`tests/fixtures/${reference.sourceFile}`);
    const hashes: Record<string, string> = {
      source: createHash('sha256').update(source).digest('hex'),
    };
    const capture = async (stage: string, pdf = false) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      let waiting = page.waitForEvent('download');
      await page
        .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
        .click();
      await (await waiting).saveAs(`${root}/${stage}.docx`);
      hashes[`${stage}.docx`] = createHash('sha256')
        .update(await fs.readFile(`${root}/${stage}.docx`))
        .digest('hex');
      if (pdf) {
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        waiting = page.waitForEvent('download');
        await page.getByRole('button', { name: 'PDF file', exact: true }).click();
        await (await waiting).saveAs(`${root}/${stage}.pdf`);
        hashes[`${stage}.pdf`] = createHash('sha256')
          .update(await fs.readFile(`${root}/${stage}.pdf`))
          .digest('hex');
      }
    };
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `Options ${name}.docx`,
      buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(page.locator('.section-page')).toHaveCount(6);
    const bodyText = await body.innerText();
    await openOptions(page);
    await page.getByRole('textbox', { name: 'Header from top (pt)', exact: true }).fill('-1');
    await expect(
      page.getByRole('button', { name: 'Apply page options', exact: true }),
    ).toBeDisabled();
    await expect(page.getByRole('alert')).toContainText('distances');
    await page.getByRole('textbox', { name: 'Header from top (pt)', exact: true }).fill('30');
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.getByRole('button', { name: 'Close dialog', exact: true }).click();
    const action = sample.action as {
      first?: number[];
      even?: boolean;
      section?: number;
      distance?: string;
      value?: number;
    };
    const targets = action.first || [action.section || 1];
    let edits = 0;
    for (const section of targets) {
      await openOptions(page);
      await page
        .getByLabel('Header/footer section', { exact: true })
        .selectOption({ label: `Section ${section}` });
      const first = page.getByRole('checkbox', { name: 'Different first page', exact: true }),
        even = page.getByRole('checkbox', { name: 'Different odd and even pages', exact: true });
      await expect(first).toBeChecked();
      await expect(
        page.getByRole('textbox', { name: 'Header from top (pt)', exact: true }),
      ).toHaveValue('18');
      if (action.first) await first.uncheck();
      if (action.even) await even.uncheck();
      if (action.distance)
        await page
          .getByRole('textbox', {
            name:
              action.distance === 'HeaderDistance'
                ? 'Header from top (pt)'
                : 'Footer from bottom (pt)',
            exact: true,
          })
          .fill(String(action.value));
      await page.getByRole('button', { name: 'Apply page options', exact: true }).click();
      await expect(page.getByRole('dialog')).toHaveCount(0);
      if (name !== 'baseline') edits++;
    }
    const check = async () => {
      await expect(page.locator('.section-page')).toHaveCount(sample.paintPages.length);
      await waitWordLayout(body);
      await expect
        .poll(async () => (await wordLineOrigins(body)).flat().map((l) => l.page))
        .toEqual(
          sample.paintPages.flatMap((p) =>
            p.lines.filter((l) => l.text.startsWith('B')).map(() => p.page),
          ),
        );
      for (const [index, p] of sample.paintPages.entries())
        await expect(
          page.locator('.section-page').nth(index).locator('.word-page-story'),
        ).toHaveText([
          p.lines
            .filter((l) => l.text.includes('header'))
            .map((l) => l.text)
            .join(''),
          p.lines
            .filter((l) => l.text.includes('footer'))
            .map((l) => l.text)
            .join(''),
        ]);
      await expect(body).toHaveText(bodyText, { useInnerText: true });
    };
    await check();
    await capture('edited', true);
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    if (!edits)
      await expect(page.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
    for (let i = 0; i < edits; i++)
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await capture('undo');
    expect(hashes['undo.docx']).toBe(hashes.source);
    for (let i = 0; i < edits; i++)
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await check();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: `Options ${name}`, exact: true }).click();
    await check();
    await capture('reloaded');
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/reloaded.docx`);
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'reloaded',
    );
    await check();
    await openOptions(page);
    await page
      .getByLabel('Header/footer section', { exact: true })
      .selectOption({ label: `Section ${targets.at(-1)}` });
    const native = sample.native.sections[targets.at(-1)! - 1];
    await expect(
      page.getByRole('checkbox', { name: 'Different first page', exact: true }),
    ).toBeChecked({ checked: native.first });
    await expect(
      page.getByRole('checkbox', { name: 'Different odd and even pages', exact: true }),
    ).toBeChecked({ checked: native.even });
    await expect(
      page.getByRole('textbox', { name: 'Header from top (pt)', exact: true }),
    ).toHaveValue(String(Math.round(native.header * 20) / 20));
    await expect(
      page.getByRole('textbox', { name: 'Footer from bottom (pt)', exact: true }),
    ).toHaveValue(String(Math.round(native.footer * 20) / 20));
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.getByRole('button', { name: 'Close dialog', exact: true }).click();
    // Reimported off states lack XML flags. Turning them back on exercises
    // new flag creation, not just undoing changes against the original XML.
    for (const section of targets) {
      await openOptions(page);
      await page
        .getByLabel('Header/footer section', { exact: true })
        .selectOption({ label: `Section ${section}` });
      await page.getByRole('checkbox', { name: 'Different first page', exact: true }).check();
      await page
        .getByRole('checkbox', { name: 'Different odd and even pages', exact: true })
        .check();
      await page.getByRole('textbox', { name: 'Header from top (pt)', exact: true }).fill('18');
      await page.getByRole('textbox', { name: 'Footer from bottom (pt)', exact: true }).fill('18');
      await page.getByRole('button', { name: 'Apply page options', exact: true }).click();
    }
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    for (let i = 0; i < edits; i++)
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await check();
    for (let i = 0; i < edits; i++)
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
    const restored = async () => {
      await expect(page.locator('.section-page')).toHaveCount(
        reference.cases.baseline.paintPages.length,
      );
      await waitWordLayout(body);
      await expect
        .poll(async () => (await wordLineOrigins(body)).flat().map((l) => l.page))
        .toEqual(
          reference.cases.baseline.paintPages.flatMap((p) =>
            p.lines.filter((l) => l.text.startsWith('B')).map(() => p.page),
          ),
        );
      for (const [index, native] of reference.cases.baseline.paintPages.entries())
        await expect(
          page.locator('.section-page').nth(index).locator('.word-page-story'),
        ).toHaveText([
          native.lines
            .filter((l) => l.text.includes('header'))
            .map((l) => l.text)
            .join(''),
          native.lines
            .filter((l) => l.text.includes('footer'))
            .map((l) => l.text)
            .join(''),
        ]);
    };
    await restored();
    await capture('restored', true);
    await page.locator('input[type=file][multiple]').setInputFiles(`${root}/restored.docx`);
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'restored',
    );
    await restored();
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify({ hashes, errors, edits, buildHash: await wordStoryBuildHash() }, null, 2),
    );
  });
