import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
import reference from './fixtures/native-word-story-spacing-grid.json' with { type: 'json' };

for (const kind of ['header', 'footer']) for (const change of ['font', 'wrap', 'nonlatin'])
  test(`Word story grid recovery ${kind} ${change}`, async ({ page }) => {
    const run = process.env.NOFFICE_STORY_GRID_RECOVERY_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Invalid recovery run');
    const row = reference.rows.find(row => row.name === `${kind}-before-3-0.25`)!;
    const name = `Grid recovery ${kind} ${change}`, root = `.local/word-story-spacing/grid-recovery/${run}/${kind}-${change}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(`tests/fixtures/word-story-spacing-grid-${row.name}.docx`);
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    expect(hash(source)).toBe(row.docxHash);
    const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
    const outputs: Record<string, string> = {};
    await page.context().grantPermissions(['local-fonts']); await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const story = page.locator(`.word-page-story[data-word-story*="${kind}"]`).first();
    const marked = story.locator('[data-word-story-grid-after]');
    await expect(marked).toHaveCount(5);
    const open = async () => {
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
      await page.getByRole('button', { name: new RegExp(`^Section 1 \u2014 Default ${kind}`) }).click();
    };
    const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    await open(); await editor.focus();
    const prefix = change === 'wrap' ? 'Wrap '.repeat(90) : '\u03a9';
    if (change === 'font') {
      await page.keyboard.press('Control+a');
      const size = page.getByRole('dialog').last().getByRole('spinbutton', { name: 'Font size', exact: true });
      await size.fill('0'); await page.keyboard.press('Tab'); await expect(size).toHaveValue('10');
      await size.fill('20'); await page.keyboard.press('Tab'); await expect(size).toHaveValue('20');
    } else { await page.keyboard.press('Control+Home'); await page.keyboard.insertText(prefix); }
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    await expect(marked).toHaveCount(0);
    await open();
    if (change !== 'font') await expect(editor.locator('p').first()).toHaveText(prefix + 'A\tB');
    else await expect(page.getByRole('dialog').last().getByRole('spinbutton', { name: 'Font size', exact: true })).toHaveValue('20');
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    const download = async (file: string, label: string) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click(); const pending = page.waitForEvent('download');
      await page.getByRole('button', { name: label, exact: true }).click(); await (await pending).saveAs(root + '/' + file);
      outputs[file] = hash(await fs.readFile(root + '/' + file));
    };
    await download('changed.docx', 'DOCX file Editable in Microsoft Word');
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(marked).toHaveCount(5);
    await page.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(marked).toHaveCount(0);
    await page.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(marked).toHaveCount(5);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click(); await expect(marked).toHaveCount(5);
    await download('recovered.docx', 'DOCX file Editable in Microsoft Word');
    expect(await fs.readFile(root + '/recovered.docx')).toEqual(source);
    await download('recovered.pdf', 'PDF file'); expect(errors).toEqual([]);
    await fs.writeFile(root + '/browser-report.json', JSON.stringify({ kind, change, sourceHash: hash(source), outputs, errors,
      buildHash: await wordStoryBuildHash(), referenceHash: hash(await fs.readFile('tests/fixtures/native-word-story-spacing-grid.json')),
      testHash: hash(await fs.readFile('tests/word-story-spacing-grid-recovery.spec.ts')) }, null, 2));
  });
