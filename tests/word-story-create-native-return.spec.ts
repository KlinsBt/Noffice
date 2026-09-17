import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-story-create.json', 'utf8'),
) as typeof import('./fixtures/native-word-story-create.json');
import { wordStoryBuildHash } from './word-story-artifacts';

test.skip(
  process.env.NOFFICE_STORY_NATIVE_RETURN !== '1',
  'Run after the independent Word section-story verifier has authored native-return.docx.',
);
for (const [kind, sample] of Object.entries(reference.cases))
  test(`Word reimports the actual ${kind} created header/footer document after a further native Word edit`, async ({
    page,
  }) => {
    const root = `.local/word-story-create/browser/${kind}`;
    const receiptBytes = await fs.readFile('.local/word-story-create/native-export-report.json');
    const receipt = JSON.parse(receiptBytes.toString('utf8').replace(/^\uFEFF/, ''));
    const source = await fs.readFile(`${root}/native-return.docx`);
    const sourceHash = createHash('sha256').update(source).digest('hex');
    expect(
      receipt.hashes.find((p: { path: string }) => p.path === `browser/${kind}/native-return.docx`)
        .sha256,
    ).toBe(sourceHash);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `Native ${kind}.docx`,
      buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const check = async () => {
      await expect(page.locator('.section-page')).toHaveCount(
        receipt.returns.find((r: { name: string }) => r.name === kind).native.pages,
      );
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      const label = { 1: 'Default', 2: 'First-page', 3: 'Even-page' }[sample.slot as 1 | 2 | 3];
      await page
        .getByRole('button', {
          name: new RegExp(
            `^Section ${sample.section} — ${label} ${sample.kind === 'Headers' ? 'header' : 'footer'}`,
          ),
        })
        .click();
      // Word's Exists flag hides inactive first/even ranges from its ordinary
      // snapshot; their retained text must still survive the native edit.
      await expect(
        page.getByRole('textbox', { name: 'Header or footer text', exact: true }).locator('p'),
      ).toHaveText([sample.text + ' native']);
      await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    };
    await check();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: `Native ${kind}`, exact: true }).click();
    await check();
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const waiting = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    await (await waiting).saveAs(`${root}/returned.docx`);
    expect(await fs.readFile(`${root}/returned.docx`)).toEqual(source);
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${root}/return-browser-report.json`,
      JSON.stringify(
        {
          sourceHash,
          buildHash: await wordStoryBuildHash(),
          nativeReceiptHash: createHash('sha256').update(receiptBytes).digest('hex'),
          errors,
        },
        null,
        2,
      ),
    );
  });
