import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

test.skip(
  process.env.NOFFICE_STORY_NATIVE_RETURN !== '1',
  'Run after the independent Word paragraph verifier has authored native-return.docx.',
);
for (const kind of ['header', 'footer'])
  test(`Word reimports the actual ${kind} paragraphs after a further native Word edit`, async ({
    page,
  }) => {
    const root = `.local/word-side-stories/paragraphs/${kind}`;
    const receiptBytes = await fs.readFile(
      '.local/word-side-stories/story-paragraph-native-report.json',
    );
    const receipt = JSON.parse(receiptBytes.toString('utf8').replace(/^\uFEFF/, ''));
    const source = await fs.readFile(`${root}/native-return.docx`);
    const sourceHash = createHash('sha256').update(source).digest('hex');
    expect(
      receipt.hashes.find(
        (p: { path: string }) => p.path === `paragraphs/${kind}/native-return.docx`,
      ).sha256,
    ).toBe(sourceHash);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `Native ${kind}.docx`,
      buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.locator('.section-page')).toHaveCount(4);
    const check = async () => {
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      await page.getByRole('button', { name: `Section 1 — Default ${kind}`, exact: true }).click();
      await expect(
        page.getByRole('textbox', { name: 'Header or footer text', exact: true }).locator('p'),
      ).toHaveText([`${kind === 'header' ? 'Header' : 'Footer'} defaultextra1`, 'extra2 native']);
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
