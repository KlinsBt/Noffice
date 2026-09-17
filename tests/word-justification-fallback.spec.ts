import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-justification-browser.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

// Synthetic mutations exercise qualification boundaries and original recovery.
// Native mode15 ignores the legacy WordPerfect flag (compatibility fixtures).
for (const variant of ['mode14', 'unknown-mode', 'word-perfect', 'font9', 'calibri', 'calibri-no-ligatures',
  'calibri12', 'calibri-bold', 'italic', 'repeated-spaces', 'paragraph-budget'])
  test(`Word justification ${variant} retains text and original export across rendering boundaries`, async ({ page }) => {
    const run = process.env.NOFFICE_JUSTIFICATION_FALLBACK || 'fallback-browser-v1';
    expect(run).toMatch(/^fallback-browser-v\d+$/);
    const row = reference.rows.find(row => row.name === 'regular-body-below')!;
    const zip = await JSZip.loadAsync(await fs.readFile(row.fixture));
    let settings = await zip.file('word/settings.xml')!.async('string');
    if (variant === 'mode14' || variant === 'unknown-mode')
      settings = settings.replace(/(w:name="compatibilityMode"[^>]*w:val=")15"/, `$1${variant === 'mode14' ? '14' : '16'}"`);
    if (variant === 'word-perfect') settings = settings.replace('<w:compat>', '<w:compat><w:wpJustification/>');
    zip.file('word/settings.xml', settings);
    let document = await zip.file('word/document.xml')!.async('string');
    const original = document.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/)![0];
    let paragraph = original, text = row.text;
    const formatting = variant === 'font9' ? '<w:sz w:val="18"/>' : variant.startsWith('calibri')
      ? `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="${variant === 'calibri12' ? 24 : 22}"/>`
        + (variant === 'calibri' ? '' : `<w14:ligatures w14:val="${variant === 'calibri-no-ligatures' ? 'none' : 'standardContextual'}"/>`)
        + (variant === 'calibri-bold' ? '<w:b/>' : '') : variant === 'italic' ? '<w:i/>' : '';
    if (formatting) paragraph = paragraph.replace('<w:t>', `<w:rPr>${formatting}</w:rPr><w:t>`);
    if (variant === 'repeated-spaces') text = row.text.replace(' ', '  ');
    if (variant === 'paragraph-budget') text = Array(30).fill(row.text).join(' ');
    paragraph = paragraph.replace(row.text, text);
    document = document.replace(original, paragraph); zip.file('word/document.xml', document);
    const input = await zip.generateAsync({ type: 'nodebuffer' });
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    const root = `.local/word-story-paragraph-layout/justification-arial-ten/${run}/${variant}`;
    await fs.mkdir(root, { recursive: true }); await fs.writeFile(root + '/input.docx', input);
    const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
    await page.context().grantPermissions(['local-fonts']); await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({ name: variant + '.docx', buffer: input,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const target = page.getByRole('textbox', { name: 'Document text', exact: true }).locator('p').first();
    await expect(target).toHaveText(text);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const qualified = variant === 'word-perfect' || variant === 'calibri';
    if (qualified) await expect(target).toHaveAttribute('data-word-justification', 'modern');
    else {
      await expect(target).not.toHaveAttribute('data-word-justification', 'modern');
      await expect(target.locator('[data-word-justification-break]')).toHaveCount(0);
    }
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
    await (await pending).saveAs(root + '/original.docx');
    expect(hash(await fs.readFile(root + '/original.docx'))).toBe(hash(input));
    await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: variant, exact: true }).click();
    await expect(target).toHaveText(text);
    if (qualified) await expect(target).toHaveAttribute('data-word-justification', 'modern');
    else await expect(target).not.toHaveAttribute('data-word-justification', 'modern');
    expect(errors).toEqual([]);
    await fs.writeFile(root + '/browser-report.json', JSON.stringify({ variant, sourceHash: row.sourceHash, inputHash: hash(input),
      errors, buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-justification-fallback.spec.ts')) }, null, 2));
  });
