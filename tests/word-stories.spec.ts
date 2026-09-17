import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

test.describe.configure({ mode: 'parallel' });
const editor = (page: Page) =>
  page.getByRole('textbox', { name: 'Header or footer text', exact: true });
async function openHeader(page: Page) {
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: 'Section 1 — Default header', exact: true }).click();
  await expect(editor(page)).toBeVisible();
}
async function exportDocx(page: Page, path: string) {
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const waiting = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
    .click();
  await (await waiting).saveAs(path);
  return fs.readFile(path);
}
for (const name of ['default', 'first-even', 'long-header', 'long-footer', 'varying-slots'])
  test(`Word ${name} header editing preserves stories through history, reload and actual export`, async ({
    page,
  }, testInfo) => {
    test.setTimeout(90000);
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    const output = testInfo.repeatEachIndex
      ? `.local/word-side-stories/repetitions/${testInfo.repeatEachIndex}/${name}`
      : `.local/word-side-stories/browser/${name}`;
    await fs.mkdir(output, { recursive: true });
    const source = await fs.readFile(`tests/fixtures/word-side-stories/${name}.docx`);
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `${name}.docx`,
      buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(body).toContainText('body01');
    const bodyText = await body.innerText();
    await openHeader(page);
    const originalText = await editor(page).innerText();
    const expectedText = originalText.replace(/^[^\r\n]*/, 'Updated header');
    await editor(page).focus();
    await page.keyboard.press('Control+End');
    await page.keyboard.insertText(' cancelled');
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await openHeader(page);
    await expect(editor(page)).toHaveText(originalText, { useInnerText: true });
    await editor(page).focus();
    await page.keyboard.press('Control+Home');
    await page.keyboard.press('Shift+End');
    await page.keyboard.insertText('Updated header');
    await expect(editor(page)).toHaveText(expectedText, { useInnerText: true });
    await page.keyboard.press('Control+Home');
    await page.keyboard.press('Shift+End');
    await page.getByRole('dialog').getByRole('button', { name: 'Bold', exact: true }).click();
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await expect(page.getByRole('dialog')).toHaveCount(0);
    await expect(body).toHaveText(bodyText, { useInnerText: true });
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    const undone = await exportDocx(page, `${output}/undo.docx`);
    expect(undone.equals(source)).toBe(true);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    const edited = await exportDocx(page, `${output}/edited.docx`);
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    await openHeader(page);
    await expect(editor(page)).toHaveText(expectedText, { useInnerText: true });
    await expect(editor(page).locator('strong')).toContainText('Updated header');
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    const reloaded = await exportDocx(page, `${output}/reloaded.docx`);
    const before = await JSZip.loadAsync(edited),
      after = await JSZip.loadAsync(reloaded);
    for (const path of Object.keys(before.files))
      if (!before.files[path].dir)
        expect(await after.file(path)!.async('uint8array'), path).toEqual(
          await before.file(path)!.async('uint8array'),
        );
    await page.locator('input[type=file][multiple]').setInputFiles(`${output}/reloaded.docx`);
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'reloaded',
    );
    await expect(body).toHaveText(bodyText, { useInnerText: true });
    await openHeader(page);
    await expect(editor(page)).toHaveText(expectedText, { useInnerText: true });
    await expect(editor(page).locator('strong')).toContainText('Updated header');
    await page.screenshot({ path: `${output}/reimported.png`, fullPage: true });
    expect(errors).toEqual([]);
    await fs.writeFile(
      `${output}/browser-report.json`,
      JSON.stringify(
        {
          originalText,
          buildHash: await wordStoryBuildHash(),
          bodyText,
          errors,
          editedText: await editor(page).innerText(),
          hashes: Object.fromEntries(
            Object.entries({ source, undo: undone, edited, reloaded }).map(([key, value]) => [
              key,
              createHash('sha256').update(value).digest('hex'),
            ]),
          ),
        },
        null,
        2,
      ),
    );
  });
