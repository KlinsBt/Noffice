import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordPlacements, waitWordLayout } from './word-pagination-helpers';

test('Word authored explicit paragraphs paginate through editing, history, reload and real exports', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1500, height: 1200 });
  const root = '.local/word-authored-pagination';
  await fs.mkdir(root, { recursive: true });
  await page.goto('/');
  await page.getByRole('button', { name: 'Start Word', exact: true }).click();
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const papers = page.locator('.section-page');
  await editor.fill('line01');
  for (let i = 2; i <= 70; i++) {
    await page.keyboard.press('Shift+Enter');
    await page.keyboard.insertText(`line${String(i).padStart(2, '0')}`);
  }
  await page.keyboard.press('Control+a');
  const fonts = page.getByRole('combobox', { name: 'Font family', exact: true });
  await fonts.selectOption('Arial');
  const size = page.getByRole('spinbutton', { name: 'Font size', exact: true });
  await size.fill('10');
  await size.press('Tab');
  const spacing = page.getByRole('combobox', { name: 'Line spacing', exact: true });
  await spacing.selectOption('custom');
  const dialog = page.getByRole('dialog', { name: 'Line spacing', exact: true });
  await dialog.getByRole('combobox', { name: 'Spacing rule', exact: true }).selectOption('exact');
  const amount = dialog.getByRole('spinbutton', { name: 'Spacing amount', exact: true });
  await amount.fill('0');
  await expect(dialog.getByRole('button', { name: 'Apply', exact: true })).toBeDisabled();
  await amount.fill('12');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(papers).toHaveCount(2);
  await expect(editor.locator('p')).toHaveCount(1);
  await waitWordLayout(editor);
  const exports: Record<string, string> = {};
  const pdfExports: Record<string, string> = {};
  const pdf = async (stage: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'PDF file', exact: true }).click();
    const bytes = await fs.readFile((await (await pending).path())!);
    expect(bytes.subarray(0, 5).toString()).toBe('%PDF-');
    await fs.writeFile(`${root}/download-${stage}.pdf`, bytes);
    pdfExports[stage] = createHash('sha256').update(bytes).digest('hex');
  };
  // A denied installed-font permission must fail without changing the document,
  // and granting it for the next explicit export must recover.
  await page.context().grantPermissions([]);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('Font access was not granted');
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  await page.context().grantPermissions(['local-fonts']);
  const capture = async (stage: string) => {
    await waitWordLayout(editor);
    const positions = await wordPlacements(editor);
    expect(positions.flat().every((c) => c.page > 0)).toBe(true);
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/browser-${stage}.docx`, bytes);
    exports[stage] = createHash('sha256').update(bytes).digest('hex');
    await page.pdf({ path: `${root}/browser-${stage}.pdf`, preferCSSPageSize: true });
    await pdf(stage);
    return positions;
  };
  const initial = await capture('source');
  await editor.click();
  await page.keyboard.press('Control+End');
  await page.keyboard.insertText(' added');
  await expect(editor).toContainText('line70 added');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor).not.toContainText(' added');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  const edited = await capture('edited');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Untitled document', exact: true }).click();
  await expect(papers).toHaveCount(2);
  const restored = await capture('restored');
  expect(restored).toEqual(initial);
  // Explicit paragraph-mark fonts now provide automatic/minimum metrics for
  // authored text. Compare each actual output independently with Word.
  await editor.click();
  await page.keyboard.press('Control+Home');
  await spacing.selectOption('1.5');
  await expect(papers).toHaveCount(2);
  const automatic = await capture('automatic');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(papers).toHaveCount(2);
  await waitWordLayout(editor);
  expect(await wordPlacements(editor)).toEqual(initial);
  await spacing.selectOption('custom');
  await dialog.getByRole('combobox', { name: 'Spacing rule', exact: true }).selectOption('atLeast');
  await amount.fill('18');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(papers).toHaveCount(2);
  const minimum = await capture('minimum');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await waitWordLayout(editor);
  expect(await wordPlacements(editor)).toEqual(initial);
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Authored reimport.docx',
    buffer: await fs.readFile(`${root}/browser-restored.docx`),
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
    'Authored reimport',
  );
  await expect(papers).toHaveCount(2);
  await waitWordLayout(editor);
  const reimported = await wordPlacements(editor);
  expect(reimported).toEqual(initial);
  await page.pdf({ path: `${root}/browser-reimported.pdf`, preferCSSPageSize: true });
  await pdf('reimported');
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        exports,
        pdfExports,
        initial,
        edited,
        restored,
        reimported,
        automatic,
        minimum,
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ),
  );
});
