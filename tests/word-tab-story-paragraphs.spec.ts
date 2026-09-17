import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

test('tabbed header paragraph split and join preserve history, reload and actual exports', async ({ page }) => {
  test.setTimeout(90000);
  const root = '.local/word-tab-story-paragraphs/browser';await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const errors: string[] = [];page.on('pageerror', (e) => errors.push(e.message));
  await page.addInitScript(() => {
    const trace: unknown[] = [];(window as unknown as { tabCaretTrace: unknown[] }).tabCaretTrace = trace;
    for (const type of ['keydown', 'keyup']) document.addEventListener(type, (event) => {
      const host = (event.target as HTMLElement).closest<HTMLElement>('[aria-label="Header or footer text"]');
      const editor = (host as (HTMLElement & { editor: import('@tiptap/core').Editor }) | null)?.editor;
      if (!editor || trace.length > 300) return;const selection = document.getSelection();
      trace.push({ event: type + ' ' + (event as KeyboardEvent).key, head: editor.state.selection.head,
        offset: editor.state.selection.$head.parentOffset, node: selection?.focusNode?.nodeName,
        text: selection?.focusNode?.textContent?.slice(0,60), domOffset: selection?.focusOffset,
        parent: selection?.focusNode?.parentElement?.outerHTML.slice(0,600),
        measured: [...host!.querySelectorAll<HTMLElement>('[data-word-tab]')].map((tab) => tab.dataset.wordTabMeasured) });
    });
  });
  await page.setViewportSize({ width: 1500, height: 1200 });await page.goto('/');
  await page.context().grantPermissions(['local-fonts']);
  await page.locator('input[type=file][multiple]').setInputFiles('tests/fixtures/word-tab-stops-stories.docx');
  await expect(page.getByRole('textbox', { name: 'Document text', exact: true })).toBeVisible();
  const open = async () => {
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
    return page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  };
  const outputs: Record<string, string> = {};
  const download = async (stage: string, format: 'docx' | 'pdf') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: format === 'docx' ? 'DOCX file Editable in Microsoft Word' : 'PDF file', exact: true }).click();
    const name = stage + '.' + format;await (await waiting).saveAs(root + '/' + name);outputs[name] = hash(await fs.readFile(root + '/' + name));
  };
  for (const stage of ['split', 'joined'] as const) {
    const story = await open();await story.focus();await page.keyboard.press('Control+Home');
    if (stage === 'split') {
      for (let i = 0; i < 4; i++) await page.keyboard.press('ArrowRight');
      await page.keyboard.press('Enter');
    } else {
      await page.keyboard.press('ArrowDown');await page.keyboard.press('Home');
      await page.keyboard.press('Backspace');
    }
    await fs.writeFile(root + '/' + stage + '-navigation-' + test.info().repeatEachIndex + '.json', JSON.stringify(await page.evaluate(() => (window as unknown as { tabCaretTrace: unknown[] }).tabCaretTrace), null, 2));
    const expected = stage === 'split' ? ['Left', '\tCenter\tRight'] : ['Left\tCenter\tRight'];
    const before = stage === 'split' ? ['Left\tCenter\tRight'] : ['Left', '\tCenter\tRight'];
    await expect(story.locator('p')).toHaveText(expected);
    await page.keyboard.press('Control+z');await expect(story.locator('p')).toHaveText(before);
    await page.keyboard.press('Control+y');await expect(story.locator('p')).toHaveText(expected);
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await expect(page.locator('.word-page-story')).toHaveCount(6);
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect((await open()).locator('p')).toHaveText(before);await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.getByRole('button', { name: 'Home', exact: true }).click();await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: 'word-tab-stops-stories', exact: true }).click();
    await expect((await open()).locator('p')).toHaveText(expected);await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await download(stage, 'docx');await download(stage, 'pdf');
  }
  for (const stage of ['split', 'joined'] as const) {
    await page.locator('input[type=file][multiple]').setInputFiles({ name: 'Returned ' + stage + '.docx', buffer: await fs.readFile(root + '/' + stage + '.docx'), mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    await expect((await open()).locator('p')).toHaveText(stage === 'split' ? ['Left', '\tCenter\tRight'] : ['Left\tCenter\tRight']);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await download(stage + '-reimported', 'docx');expect(outputs[stage + '-reimported.docx']).toBe(outputs[stage + '.docx']);
    await download(stage + '-reimported', 'pdf');
  }
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(await fs.readFile('tests/fixtures/word-tab-stops-stories.docx')), outputs, errors, buildHash: await wordStoryBuildHash() }, null, 2));
});

for (const stage of ['split', 'joined'] as const)
  test(`native ${stage} header further edit returns through browser history and reload`, async ({ page }) => {
    test.skip(process.env.NOFFICE_TAB_NATIVE_RETURN !== '1', 'Requires separately captured installed-Word files.');
    const root = '.local/word-tab-story-paragraphs', folder = root + '/' + (process.env.NOFFICE_TAB_PARAGRAPH_RUN || 'export-native');
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    const receipt = JSON.parse((await fs.readFile(folder + '/native-report.json', 'utf8')).replace(/^\uFEFF/, ''));
    const row = receipt.rows.find((row: { stage: string }) => row.stage === stage);
    const input = await fs.readFile(folder + '/' + stage + '-returned.docx');expect(hash(input)).toBe(row.returnedHash);
    const errors: string[] = [];page.on('pageerror', (e) => errors.push(e.message));
    const name = 'Native paragraph ' + stage;await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: input, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    await expect(page.getByRole('textbox', { name: 'Document text', exact: true })).toBeVisible();
    const open = async () => {
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
      return page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    };
    const expected = stage === 'split' ? ['Later', '\tCenter\tRight'] : ['Later\tCenter\tRight'];
    const story = await open();await expect(story.locator('p')).toHaveText(expected);
    await story.focus();await page.keyboard.press('Control+Home');await page.keyboard.type('Reviewed ');
    await expect(story.locator('p')).toHaveText(['Reviewed ' + expected[0], ...expected.slice(1)]);
    await page.keyboard.press('Control+z');await expect(story.locator('p')).toHaveText(expected);
    await page.keyboard.press('Control+y');await expect(story.locator('p')).toHaveText(['Reviewed ' + expected[0], ...expected.slice(1)]);
    await page.keyboard.press('Control+z');await expect(story.locator('p')).toHaveText(expected);
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();await page.getByRole('button', { name, exact: true }).click();
    await expect((await open()).locator('p')).toHaveText(expected);await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await page.getByRole('button', { name: 'Export', exact: true }).click();const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
    const target = root + '/browser/' + stage + '-returned.docx';await (await waiting).saveAs(target);
    expect(await fs.readFile(target)).toEqual(input);expect(errors).toEqual([]);
    await fs.writeFile(root + '/browser/' + stage + '-return-browser-report.json', JSON.stringify({ sourceHash: row.returnedHash,
      nativeReceiptHash: hash(await fs.readFile(folder + '/native-report.json')), errors, buildHash: await wordStoryBuildHash() }, null, 2));
  });
