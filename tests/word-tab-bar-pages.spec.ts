import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

for (const name of ['continuous-dimensions', 'columns-balanced', 'columns-overflow', 'next-page'])
  test(`${name} bar stops follow body fragments through editing, reload and print`, async ({ page }) => {
    test.setTimeout(90000);
    const root = '.local/word-tab-bars/pages/' + name;await fs.mkdir(root, { recursive: true });
    const zip = await JSZip.loadAsync(await fs.readFile(`tests/fixtures/word-section-flow-${name === 'next-page' ? 'continuous-dimensions' : name}.docx`));
    let xml = await zip.file('word/document.xml')!.async('string');
    if (name === 'next-page') xml = xml.replaceAll('w:val="continuous"', 'w:val="nextPage"');
    zip.file('word/document.xml', xml.replaceAll('</w:pPr>', '<w:tabs><w:tab w:val="bar" w:pos="1440"/></w:tabs></w:pPr>'));
    const bytes = await zip.generateAsync({ type: 'nodebuffer' });await fs.writeFile(root + '/input.docx', bytes);
    await page.setViewportSize({ width: 1500, height: 1200 });await page.goto('/');await page.context().grantPermissions(['local-fonts']);
    await page.locator('input[type=file][multiple]').setInputFiles({ name: `bar-${name}.docx`, buffer: bytes, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });await expect(body).toBeVisible();
    await expect(page.locator('.section-page').first()).toBeVisible();
    await expect(page.locator('.section-page .word-tab-bar').first()).toBeVisible();
    const sourceText = await body.innerText();
    const outputs: Record<string, string> = {};const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
    const download = async (stage: string, type: 'docx' | 'pdf') => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();const waiting = page.waitForEvent('download');
      await page.getByRole('button', { name: type === 'docx' ? 'DOCX file Editable in Microsoft Word' : 'PDF file', exact: true }).click();
      const file = `${stage}.${type}`;await (await waiting).saveAs(root + '/' + file);outputs[file] = hash(await fs.readFile(root + '/' + file));
    };
    await download('source', 'docx');expect(outputs['source.docx']).toBe(hash(bytes));await download('source', 'pdf');
    await body.focus();await page.keyboard.press('Control+Home');await page.keyboard.insertText('Edited ');
    await expect(body).toContainText('Edited');await page.keyboard.press('Control+z');expect(await body.innerText()).toBe(sourceText);await page.keyboard.press('Control+y');await expect(body).toContainText('Edited');
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();await page.getByRole('button', { name: `bar-${name}`, exact: true }).click();
    await expect(body).toContainText('Edited');await expect(page.locator('.section-page .word-tab-bar').first()).toBeVisible();
    await download('edited', 'docx');await download('edited', 'pdf');
    await page.emulateMedia({ media: 'print' });await page.pdf({ path: root + '/css-print.pdf', preferCSSPageSize: true, printBackground: true });await page.emulateMedia({ media: 'screen' });
    const paint = await page.locator('.section-page .word-tab-bar').evaluateAll((bars) => bars.map((bar) => {
      const r = bar.getBoundingClientRect(), p = bar.closest('.section-page')!.getBoundingClientRect();return { x: (r.left-p.left)*.75, y: (r.top-p.top)*.75, width: r.width*.75, height: r.height*.75 };
    }));
    await page.getByRole('button', { name: 'Open', exact: true }).click();
    await page.locator('input[type=file][multiple]').setInputFiles({ name: `bar-${name}-returned.docx`, buffer: await fs.readFile(root + '/edited.docx'), mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    await expect(body).toContainText('Edited');await expect(page.locator('.section-page .word-tab-bar').first()).toBeVisible();
    await download('reimported', 'docx');expect(outputs['reimported.docx']).toBe(outputs['edited.docx']);await download('reimported', 'pdf');
    await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(bytes), outputs, paint, buildHash: await wordStoryBuildHash() }, null, 2));
  });

for (const name of ['continuous-dimensions', 'columns-balanced', 'columns-overflow', 'next-page'])
  test(`native ${name} body bars return through keyboard history and reload`, async ({ page }) => {
    test.skip(process.env.NOFFICE_TAB_BAR_NATIVE_RETURN !== '1', 'Requires independently edited installed-Word outputs');
    const run = process.env.NOFFICE_TAB_BAR_NATIVE_RUN || 'native-v1';
    const folder = '.local/word-tab-bars/pages-' + run, root = '.local/word-tab-bars/pages/' + name;
    const bytes = await fs.readFile(`${folder}/${name}-returned.docx`);const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
    await page.setViewportSize({ width: 1500, height: 1200 });await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({ name: `native-bar-${name}.docx`, buffer: bytes, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });await expect(body).toContainText('Native');
    await body.focus();await page.keyboard.press('Control+Home');await page.keyboard.insertText('Reviewed ');await expect(body).toContainText('Reviewed Native');
    await page.keyboard.press('Control+z');await expect(body).not.toContainText('Reviewed');await page.keyboard.press('Control+y');await expect(body).toContainText('Reviewed Native');await page.keyboard.press('Control+z');
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();await page.getByRole('button', { name: `native-bar-${name}`, exact: true }).click();
    await expect(body).toContainText('Native');await expect(page.locator('.section-page .word-tab-bar').first()).toBeVisible();
    await page.getByRole('button', { name: 'Export', exact: true }).click();const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();await (await waiting).saveAs(root + '/native-return.docx');expect(await fs.readFile(root + '/native-return.docx')).toEqual(bytes);
    await fs.writeFile(root + '/native-return-report.json', JSON.stringify({ sourceHash: hash(bytes), nativeReceiptHash: hash(await fs.readFile(folder + '/native-report.json')), buildHash: await wordStoryBuildHash() }, null, 2));
  });
