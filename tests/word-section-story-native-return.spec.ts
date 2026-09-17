import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

test.skip(
  process.env.NOFFICE_STORY_NATIVE_RETURN !== '1',
  'Run after the independent Word section-story verifier has authored native-return.docx.',
);
for (const kind of [
  'linked-next',
  'independent-next',
  'linked-width',
  'first-even',
  'continuous-linked',
  'continuous-independent',
  'odd-start',
  'even-start',
])
  test(`Word reimports the actual ${kind} section header after a further native Word edit`, async ({
    page,
  }) => {
    const root = `.local/word-section-stories/browser/${kind}`;
    const receiptBytes = await fs.readFile('.local/word-section-stories/native-export-report.json');
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
    await expect(
      page.locator('.word-page-story').filter({ hasText: 'Changed second header native' }).first(),
    ).toBeVisible();
    const check = async () => {
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      await page.getByRole('button', { name: /^Section 2 \u2014 Default header/ }).click();
      await expect(
        page.getByRole('textbox', { name: 'Header or footer text', exact: true }).locator('p'),
      ).toHaveText(['Changed second header native']);
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
