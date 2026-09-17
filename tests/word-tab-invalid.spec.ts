import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

for (const name of ['empty-default', 'empty-decimal', 'invalid-clear', 'invalid-bar'])
  test(`${name} tab definitions preserve the original and reject an incomplete PDF`, async ({ page }) => {
    const root = '.local/word-tab-invalid/' + name;await fs.mkdir(root, { recursive: true });
    const zip = await JSZip.loadAsync(await fs.readFile('tests/fixtures/word-tab-stops-stories.docx'));
    if (name.startsWith('empty-')) {
      const tag = name === 'empty-default' ? 'defaultTabStop' : 'decimalSymbol';
      const path = 'word/settings.xml', source = await zip.file(path)!.async('string');
      const changed = source.replace(new RegExp(`<w:${tag}\\b[^>]*/>`), `<w:${tag}/>`);expect(changed).not.toBe(source);zip.file(path, changed);
    } else {
      const path = 'word/header2.xml', source = await zip.file(path)!.async('string');
      const tabs = name === 'invalid-clear' ? '<w:tab w:val="clear" w:pos="bad"/>' : '<w:tab w:val="bar" w:pos="31681"/>';
      let changed = source.replace('</w:pPr>', '<w:tabs>' + tabs + '</w:tabs></w:pPr>');expect(changed).not.toBe(source);
      if (name === 'invalid-bar') changed = changed.replace(/<w:tab\s*\/>/g, '');zip.file(path, changed);
    }
    const input = await zip.generateAsync({ type: 'nodebuffer' });await fs.writeFile(root + '/input.docx', input);
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    const errors: string[] = [];page.on('pageerror', (e) => errors.push(e.message));
    await page.goto('/');await page.context().grantPermissions(['local-fonts']);
    await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: input, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    await expect(page.getByRole('textbox', { name: 'Document text', exact: true })).toBeVisible();
    await page.getByRole('button', { name: 'Export', exact: true }).click();let waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();await (await waiting).saveAs(root + '/preserved.docx');
    expect(await fs.readFile(root + '/preserved.docx')).toEqual(input);
    let downloads = 0;page.on('download', () => downloads++);
    await page.getByRole('button', { name: 'Export', exact: true }).click();await page.getByRole('button', { name: 'PDF file', exact: true }).click();
    const failure = "PDF export cannot lay out the document's unsupported tab definitions. The original DOCX is preserved.";
    await expect(page.getByRole('alert')).toContainText('Export failed: ' + failure);expect(downloads).toBe(0);
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await body.focus();await page.keyboard.press('Control+Home');await page.keyboard.type('Local ');
    await expect(body.locator('p').first()).toHaveText(/^Local line01/);
    await page.keyboard.press('Control+z');await expect(body.locator('p').first()).toHaveText(/^line01/);
    await page.keyboard.press('Control+y');await expect(body.locator('p').first()).toHaveText(/^Local line01/);
    await page.keyboard.press('Control+z');await expect(body.locator('p').first()).toHaveText(/^line01/);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();await page.getByRole('button', { name, exact: true }).click();
    await expect(body.locator('p').first()).toHaveText(/^line01/);
    await page.getByRole('button', { name: 'Export', exact: true }).click();waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();await (await waiting).saveAs(root + '/after-history.docx');
    expect(await fs.readFile(root + '/after-history.docx')).toEqual(input);
    await page.getByRole('button', { name: 'Export', exact: true }).click();waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: /^Noffice backup/ }).click();await (await waiting).saveAs(root + '/recovery.noffice');
    const backup = JSON.parse(await fs.readFile(root + '/recovery.noffice', 'utf8'));expect(Buffer.from(backup.original.base64, 'base64')).toEqual(input);
    expect(errors).toEqual([]);
    await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(input), preservedHash: hash(await fs.readFile(root + '/preserved.docx')),
      failure, errors, buildHash: await wordStoryBuildHash() }, null, 2));
  });
