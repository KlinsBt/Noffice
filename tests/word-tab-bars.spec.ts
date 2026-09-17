import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

for (const name of ['plain', 'hard', 'negative', 'tabs'])
  test(`${name} header bar stops survive editing, history, reload and exports`, async ({ page }) => {
    test.setTimeout(90000);
    const root = '.local/word-tab-bars/browser/' + name;await fs.mkdir(root, { recursive: true });
    const zip = await JSZip.loadAsync(await fs.readFile('tests/fixtures/word-tab-stops-stories.docx'));
    const path = 'word/header2.xml', source = await zip.file(path)!.async('string');
    let header = source.replace('</w:pPr>', `<w:tabs><w:tab w:val="bar" w:pos="${name === 'negative' ? -360 : 1440}"/></w:tabs></w:pPr>`);
    if (name !== 'tabs') header = header.replace(/<w:tab\s*\/>/g, name === 'hard' ? '<w:br/>' : '');
    expect(header).not.toBe(source);zip.file(path, header);
    const input = await zip.generateAsync({ type: 'nodebuffer' });await fs.writeFile(root + '/input.docx', input);
    const errors: string[] = [];page.on('pageerror', (error) => errors.push(error.message));
    const hash = (value: Buffer) => createHash('sha256').update(value).digest('hex');
    const outputs: Record<string, string> = {};
    await page.setViewportSize({ width: 1500, height: 1200 });await page.goto('/');await page.context().grantPermissions(['local-fonts']);
    await page.locator('input[type=file][multiple]').setInputFiles({ name: `bar-${name}.docx`, buffer: input, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    await expect(page.locator('.section-page')).toHaveCount(3);
    const bars = page.locator('.word-page-story [data-word-bar-paint=true]');await expect(bars).toHaveCount(3);
    const paint = await bars.first().evaluate((paragraph) => {
      const style = getComputedStyle(paragraph, '::before');
      return { position: style.position, shadow: style.boxShadow, width: parseFloat(style.width), height: parseFloat(style.height) };
    });
    expect(paint.position).toBe('absolute');expect(paint.width).toBeGreaterThan(0);expect(paint.height).toBeGreaterThan(0);expect(paint.shadow).not.toBe('none');
    const download = async (stage: string, type: 'docx' | 'pdf') => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();const waiting = page.waitForEvent('download');
      await page.getByRole('button', { name: type === 'docx' ? 'DOCX file Editable in Microsoft Word' : 'PDF file', exact: true }).click();
      await expect(page.getByRole('alert').filter({ hasText: 'Export failed:' })).toHaveCount(0);
      const file = `${stage}.${type}`;await (await waiting).saveAs(root + '/' + file);outputs[file] = hash(await fs.readFile(root + '/' + file));
    };
    await download('source', 'docx');expect(outputs['source.docx']).toBe(hash(input));await download('source', 'pdf');
    await page.getByRole('button', { name: 'Insert', exact: true }).click();await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
    const story = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    await story.focus();await page.keyboard.press('Control+Home');for (let i = 0; i < 4; i++) await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.insertText('Edited');await expect(story).toContainText('Edited');
    await page.keyboard.press('Control+z');await expect(story).toContainText('Left');await page.keyboard.press('Control+y');await expect(story).toContainText('Edited');
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    const changed = page.locator('.word-page-story').filter({ hasText: 'Edited' });await expect(changed).toHaveCount(3);
    await page.getByRole('button', { name: 'Home', exact: true }).click();await page.getByRole('button', { name: 'Undo', exact: true }).click();await expect(changed).toHaveCount(0);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();await expect(changed).toHaveCount(3);await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();await page.getByRole('button', { name: `bar-${name}`, exact: true }).click();
    await expect(changed).toHaveCount(3);await expect(bars).toHaveCount(3);await download('edited', 'docx');await download('edited', 'pdf');
    await page.getByRole('button', { name: 'Open', exact: true }).click();
    await page.locator('input[type=file][multiple]').setInputFiles({ name: `bar-${name}-returned.docx`, buffer: await fs.readFile(root + '/edited.docx'), mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    await expect(changed).toHaveCount(3);await expect(bars).toHaveCount(3);
    await download('reimported', 'docx');expect(outputs['reimported.docx']).toBe(outputs['edited.docx']);await download('reimported', 'pdf');
    expect(errors).toEqual([]);await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(input), outputs, errors, paint, buildHash: await wordStoryBuildHash() }, null, 2));
  });

for (const name of ['plain', 'hard', 'negative', 'tabs'])
  test(`native ${name} bar header returns through browser editing and reload`, async ({ page }) => {
    test.skip(process.env.NOFFICE_TAB_BAR_NATIVE_RETURN !== '1', 'Requires independently edited installed-Word outputs');
    const run = process.env.NOFFICE_TAB_BAR_NATIVE_RUN || 'native-v1';
    const folder = '.local/word-tab-bars/' + run, root = '.local/word-tab-bars/browser/' + name;
    const bytes = await fs.readFile(`${folder}/${name}-returned.docx`);
    const hash = (value: Buffer) => createHash('sha256').update(value).digest('hex');
    await page.setViewportSize({ width: 1500, height: 1200 });await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({ name: `native-bar-${name}.docx`, buffer: bytes, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    await expect(page.locator('.word-page-story').filter({ hasText: 'Native' })).toHaveCount(3);
    await page.getByRole('button', { name: 'Insert', exact: true }).click();await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
    const story = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    await story.focus();await page.keyboard.press('Control+Home');await page.keyboard.insertText('Reviewed ');await expect(story).toContainText('Reviewed Native');
    await page.keyboard.press('Control+z');await expect(story).not.toContainText('Reviewed');await page.keyboard.press('Control+y');await expect(story).toContainText('Reviewed Native');
    await page.keyboard.press('Control+z');await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();await page.getByRole('button', { name: `native-bar-${name}`, exact: true }).click();
    await expect(page.locator('.word-page-story [data-word-bar-paint=true]')).toHaveCount(3);
    await page.getByRole('button', { name: 'Export', exact: true }).click();const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();await (await waiting).saveAs(root + '/native-return.docx');
    expect(await fs.readFile(root + '/native-return.docx')).toEqual(bytes);
    await fs.writeFile(root + '/native-return-report.json', JSON.stringify({ sourceHash: hash(bytes), nativeReceiptHash: hash(await fs.readFile(folder + '/native-report.json')), buildHash: await wordStoryBuildHash() }, null, 2));
  });
