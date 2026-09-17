import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { wordStoryBuildHash } from './word-story-artifacts';

test('Fresh Word paragraph spacing: line and Auto modes, contextual toggle, history, actual export and reimport', async ({ page }) => {
  test.setTimeout(90000);
  const root = '.local/word-paragraph-spacing-modes/fresh-browser-v1'; await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const outputs: Record<string, string> = {}, errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.getByRole('button', { name: 'Start Word', exact: true }).click();
  const body = page.getByRole('textbox', { name: 'Document text', exact: true }); await body.focus();
  for (const [index, text] of [...'ABCDE'].entries()) { if (index) await page.keyboard.press('Enter'); await page.keyboard.insertText(text); }
  await page.keyboard.press('Control+a');
  await page.getByRole('combobox', { name: 'Font family', exact: true }).selectOption('Arial');
  const size = page.getByRole('spinbutton', { name: 'Font size', exact: true }); await size.fill('10'); await size.press('Tab');
  await page.getByRole('combobox', { name: 'Line spacing', exact: true }).selectOption('custom');
  const dialog = page.getByRole('dialog', { name: 'Line spacing', exact: true });
  await dialog.getByRole('combobox', { name: 'Spacing rule', exact: true }).selectOption('exact');
  await dialog.getByRole('spinbutton', { name: 'Spacing amount', exact: true }).fill('40');
  await dialog.getByRole('button', { name: 'Apply', exact: true }).click();
  await page.getByRole('button', { name: 'Layout', exact: true }).click();
  const before = page.getByRole('spinbutton', { name: 'Before paragraph', exact: true });
  const after = page.getByRole('spinbutton', { name: 'After paragraph', exact: true });
  await before.fill('0'); await before.press('Enter'); await after.fill('0'); await after.press('Enter');
  const paragraphs = body.locator(':scope > p'); await expect(paragraphs).toHaveCount(5); await paragraphs.nth(2).click();
  await page.getByLabel('Before paragraph unit', { exact: true }).selectOption('lines'); await before.fill('1'); await before.press('Enter');
  await page.getByLabel('After paragraph unit', { exact: true }).selectOption('auto');
  const contextual = page.getByRole('checkbox', { name: "Don't add space between paragraphs of the same style", exact: true });
  const advances = async (expected: number[]) => {
    await expect(paragraphs).toHaveCount(5);
    await expect.poll(async () => {
      const ys = await paragraphs.evaluateAll(ps => ps.map(p => p.getBoundingClientRect().top * .75));
      return Math.max(...ys.slice(1).map((y, i) => Math.abs(y - ys[i] - expected[i])));
    }).toBeLessThanOrEqual(.15);
  };
  await contextual.check(); await advances([40, 40, 40, 40]);
  await contextual.uncheck(); await advances([40, 52, 54, 40]);
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); await advances([40, 40, 40, 40]);
  await page.getByRole('button', { name: 'Redo', exact: true }).click(); await advances([40, 52, 54, 40]);
  const download = async (name: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download'); await page.getByRole('button', { name: label, exact: true }).click();
    await (await pending).saveAs(root + '/' + name); outputs[name] = hash(await fs.readFile(root + '/' + name));
  };
  await download('fresh.docx', 'DOCX file Editable in Microsoft Word'); await download('fresh.pdf');
  const zip = await JSZip.loadAsync(await fs.readFile(root + '/fresh.docx'));
  const xml = await zip.file('word/document.xml')!.async('string'); expect(xml).toContain('w:beforeLines="100"'); expect(xml).toContain('w:afterAutospacing="1"');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const title = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click(); await page.getByRole('button', { name: title, exact: true }).click();
  await advances([40, 52, 54, 40]); await download('reloaded.pdf');
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles({ name: 'Fresh spacing reimport.docx', buffer: await fs.readFile(root + '/fresh.docx'), mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await advances([40, 52, 54, 40]); await download('reimported.pdf'); expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ outputs, errors, buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-paragraph-spacing-fresh.spec.ts')) }, null, 2));
});
