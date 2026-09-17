import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

export function storyNativeReturnTests(
  cases: { name: string; text: string; kind: 'header' | 'footer'; slot: number; pages: number }[],
  root: string,
  receiptPath: string,
) {
  test.skip(
    process.env.NOFFICE_STORY_NATIVE_RETURN !== '1',
    'Run after native-edited.docx has been authored by the independent Word verifier.',
  );
  for (const sample of cases)
    test(`Word reimports ${sample.name} after a further native edit and preserves its exported bytes`, async ({
      page,
    }) => {
      const folder = `${root}/${sample.name}`;
      const receiptBytes = await fs.readFile(receiptPath);
      const receipt = JSON.parse(receiptBytes.toString('utf8').replace(/^\uFEFF/, ''));
      const source = await fs.readFile(`${folder}/native-edited.docx`);
      const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
      const sourceHash = hash(source);
      expect(receipt.rows.find((r: { name: string }) => r.name === sample.name).returnHash).toBe(
        sourceHash,
      );
      const errors: string[] = [];
      page.on('pageerror', (e) => errors.push(e.message));
      await page.goto('/');
      await page
        .locator('input[type=file][multiple]')
        .setInputFiles({
          name: `Native ${sample.name}.docx`,
          buffer: source,
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });
      const check = async () => {
        await expect(page.locator('.section-page')).toHaveCount(sample.pages);
        await page.getByRole('button', { name: 'Insert', exact: true }).click();
        await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
        const label = ({ 1: 'Default', 2: 'First-page', 3: 'Even-page' } as const)[
          sample.slot as 1 | 2 | 3
        ];
        await page
          .getByRole('button', { name: new RegExp(`^Section 1 \\u2014 ${label} ${sample.kind}`) })
          .click();
        await expect(
          page.getByRole('textbox', { name: 'Header or footer text', exact: true }).locator('p'),
        ).toHaveText([sample.text + ' native']);
        await page.getByRole('button', { name: 'Cancel', exact: true }).click();
      };
      await check();
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      await page.reload();
      await page.getByRole('button', { name: 'Recent files', exact: true }).click();
      await page.getByRole('button', { name: `Native ${sample.name}`, exact: true }).click();
      await check();
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const waiting = page.waitForEvent('download');
      await page
        .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
        .click();
      await (await waiting).saveAs(`${folder}/returned.docx`);
      expect(await fs.readFile(`${folder}/returned.docx`)).toEqual(source);
      expect(errors).toEqual([]);
      await fs.writeFile(
        `${folder}/return-browser-report.json`,
        JSON.stringify(
          {
            sourceHash,
            buildHash: await wordStoryBuildHash(),
            nativeReceiptHash: hash(receiptBytes),
            errors,
          },
          null,
          2,
        ),
      );
    });
}
