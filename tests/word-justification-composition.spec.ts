import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import arial from './fixtures/native-word-justification-browser.json' with { type: 'json' };
import calibri from './fixtures/native-word-justification-calibri-browser.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
const profile = process.env.NOFFICE_JUSTIFICATION_PROFILE || 'arial-ten';
if (!['arial-ten', 'calibri-eleven'].includes(profile)) throw Error('Invalid justification profile');
const reference = profile === 'calibri-eleven' ? calibri : arial;

for (const kind of ['body', 'header', 'footer']) test(`Word justification ${kind}: browser composition across a visual break`, async ({ page }) => {
  const row = reference.rows.find(row => row.name === `regular-${kind}-below`)!;
  const run = process.env.NOFFICE_JUSTIFICATION_COMPOSITION || 'composition-browser-v1';
  expect(run).toMatch(/^composition-browser-v\d+$/);
  const root = `.local/word-story-paragraph-layout/justification-${profile}/${run}/${kind}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles(row.fixture);
  await expect(page.locator('.section-page')).toHaveCount(row.pages);
  let editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  if (kind !== 'body') {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: new RegExp('^Section 1 \\u2014 Default ' + kind) }).click();
    editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  }
  const paragraph = editor.locator('p').nth(kind === 'body' ? 0 : 2);
  const toolbar = kind === 'body' ? page : page.getByRole('dialog');
  await expect(paragraph).toHaveAttribute('data-word-justification', 'modern');
  const start = async () => {
    await editor.focus(); await page.keyboard.press('Control+Home');
    if (kind !== 'body') {
      await page.keyboard.press('Control+ArrowDown'); await page.keyboard.press('Control+ArrowDown');
      await page.keyboard.press('Home');
    }
  };
  const session = await page.context().newCDPSession(page);
  await start();
  await session.send('Input.imeSetComposition', { text: 'ni', selectionStart: 2, selectionEnd: 2 });
  await expect(paragraph).toHaveText('ni' + row.text);
  await session.send('Input.imeSetComposition', { text: '\u65e5\u672c', selectionStart: 2, selectionEnd: 2 });
  await expect(paragraph).toHaveText('\u65e5\u672c' + row.text);
  await session.send('Input.imeSetComposition', { text: '', selectionStart: 0, selectionEnd: 0 });
  await expect(paragraph).toHaveText(row.text);
  await expect(paragraph).toHaveAttribute('data-word-justification', 'modern');
  await start(); const selectedLength = row.ends[0] + 5;
  for (let i = 0; i < selectedLength; i++) await page.keyboard.press('Shift+ArrowRight');
  for (const text of ['Rep', 'Replace']) {
    await session.send('Input.imeSetComposition', { text, selectionStart: text.length, selectionEnd: text.length });
    await expect(paragraph).toHaveText(text + row.text.slice(selectedLength));
  }
  await session.send('Input.insertText', { text: 'Replacement' });
  const finalText = 'Replacement' + row.text.slice(selectedLength);
  await expect(paragraph).toHaveText(finalText);
  await expect(paragraph).toHaveAttribute('data-word-justification', 'modern');
  await toolbar.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(paragraph).toHaveText(row.text);
  await toolbar.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(paragraph).toHaveText(finalText);
  if (kind !== 'body') await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
  await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click();
  const savedParagraph = kind === 'body' ? paragraph : page.locator('.word-story-measure-host p').filter({ hasText: finalText }).first();
  await expect(savedParagraph).toHaveText(finalText);
  const outputs: Record<string, string> = {};
  for (const [label, file] of [['DOCX file Editable in Microsoft Word', 'edited.docx'], ['PDF file', 'edited.pdf']]) {
    await page.getByRole('button', { name: 'Export', exact: true }).click(); const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click(); await (await waiting).saveAs(root + '/' + file);
    outputs[file] = hash(await fs.readFile(root + '/' + file));
  }
  await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/edited.docx');
  await expect(savedParagraph).toHaveText(finalText); expect(errors).toEqual([]);
  await session.detach();
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ name: row.name, kind, sourceHash: row.sourceHash, errors,
    outputs, actions: { selectedLength, replacement: 'Replacement' }, finalText,
    inputMethod: 'Chromium CDP composition candidate/update/cancel/commit; native OS/Word IME not certified',
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-justification-composition.spec.ts')) }, null, 2));
});
