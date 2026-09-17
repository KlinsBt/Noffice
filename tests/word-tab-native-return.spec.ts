import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

test.skip(process.env.NOFFICE_TAB_NATIVE_RETURN !== '1', 'Requires the separately captured installed-Word return files.');

for (const kind of ['body', 'stories'] as const)
  test(`native-edited ${kind} tabs preserve further editing, history, reload and original return bytes`, async ({ page }) => {
    const root = '.local/word-tab-stops-v2', folder = root + '/' + (process.env.NOFFICE_TAB_NATIVE_RUN || 'native-current-v2');
    const receipt = JSON.parse((await fs.readFile(folder + '/native-report.json', 'utf8')).replace(/^\uFEFF/, ''));
    const row = receipt.returns.find((r: { kind: string }) => r.kind === kind);
    const input = await fs.readFile(folder + '/' + kind + '-returned.docx');
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    expect(hash(input)).toBe(row.returnHash);
    const errors: string[] = []; page.on('pageerror', (e) => errors.push(e.message));
    await page.setViewportSize({ width: 1500, height: 1200 }); await page.goto('/');
    const name = 'Native tabs ' + kind;
    await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: input,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(body).toBeVisible();
    const open = async () => {
      if (kind === 'body') return body;
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
      return page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    };
    const expected = kind === 'body' ? 'A\tB\tC\tN' : 'Edited\tCenter\tReturned';
    const editor = await open(); await expect(editor.locator('p').first()).toHaveText(expected);
    await editor.focus();
    if (kind === 'body') await page.keyboard.press('Control+Home');
    await page.keyboard.press('End'); await page.keyboard.press('Shift+Home');
    const changed = 'Changed\t' + expected;
    await page.keyboard.insertText(changed);
    await expect(editor.locator('p').first()).toHaveText(changed);
    await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText(expected);
    await page.keyboard.press('Control+y'); await expect(editor.locator('p').first()).toHaveText(changed);
    await page.keyboard.press('Control+z'); await expect(editor.locator('p').first()).toHaveText(expected);
    if (kind === 'stories') await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    await expect((await open()).locator('p').first()).toHaveText(expected);
    if (kind === 'stories') await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
    const target = root + '/browser/' + kind;
    await (await waiting).saveAs(target + '/returned.docx');
    expect(await fs.readFile(target + '/returned.docx')).toEqual(input);
    expect(errors).toEqual([]);
    await fs.writeFile(target + '/return-browser-report.json', JSON.stringify({
      sourceHash: row.returnHash, nativeReceiptHash: hash(await fs.readFile(folder + '/native-report.json')),
      errors, buildHash: await wordStoryBuildHash(),
    }, null, 2));
  });
