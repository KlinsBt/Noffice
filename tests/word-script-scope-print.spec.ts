import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import native from './fixtures/native-word-script-scopes.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of native.rows) for (const stage of ['formatted', 'typed'])
  test(`Word script scope print ${row.name} ${stage}: screen, direct PDF and browser print`, async ({ page }) => {
    test.skip(process.env.NOFFICE_SCRIPT_SCOPE_PRINT !== '1', 'Requires verified script-scope authoring downloads');
    const sourceRun = process.env.NOFFICE_SCRIPT_SCOPE_SOURCE_RUN || 'browser-v6';
    const run = process.env.NOFFICE_SCRIPT_SCOPE_PRINT_RUN || 'print-v1';
    if (!/^browser-v\d+$/.test(sourceRun) || !/^print-v\d+$/.test(run)) throw Error('Invalid evidence folder');
    const inputRoot = `.local/word-script-scopes/${sourceRun}/${row.name}`;
    const root = `${inputRoot}/${run}/${stage}`; await fs.mkdir(root, { recursive: true });
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    const input = await fs.readFile(`${inputRoot}/${stage}.docx`);
    const prior = JSON.parse((await fs.readFile(inputRoot + '/browser-report.json')).toString());
    expect(hash(input)).toBe(prior.outputs[stage + '.docx']);
    const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
    await page.setViewportSize({ width: 1500, height: 1200 });
    await page.context().grantPermissions(['local-fonts']); await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles(`${inputRoot}/${stage}.docx`);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    await expect(editor).toBeVisible(); await page.evaluate(() => document.fonts.ready);
    const paragraph = row.kind === 'Body' ? editor.locator('p').first()
      : page.locator(`.word-page-story[data-word-story*="${row.kind === 'Headers' ? 'header' : 'footer'}"]`).first().locator('p').first();
    await expect(paragraph).toBeVisible();
    const expected = (stage === 'formatted' ? row.formatted : row.typed).map(c => c.text).join('').replace(/\r$/, '');
    await expect(paragraph).toHaveText(expected);
    const model = () => editor.evaluate(element => (element as HTMLElement & { editor: { getJSON(): unknown } }).editor.getJSON());
    const before = await model(), paint = await paragraph.evaluate(p => ({
      scriptLayout: p.getAttribute('data-word-script-layout'),
      scripts: [...p.querySelectorAll('sup,sub')].map(el => ({ text: el.textContent, size: getComputedStyle(el).fontSize })),
      tabs: [...p.querySelectorAll('[data-word-tab]')].map(el => ({ html: el.outerHTML, size: getComputedStyle(el).fontSize })),
    }));
    await paragraph.scrollIntoViewIfNeeded(); await paragraph.screenshot({ path: root + '/screen.png' });
    await page.emulateMedia({ media: 'print' });
    await page.pdf({ path: root + '/browser-print.pdf', preferCSSPageSize: true, printBackground: true });
    await page.emulateMedia({ media: 'screen' }); expect(await model()).toEqual(before);
    const outputs: Record<string, string> = {};
    for (const name of ['screen.png', 'browser-print.pdf']) outputs[name] = hash(await fs.readFile(root + '/' + name));
    let failure: string | null = null;
    try {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download'); pending.catch(() => {});
      await page.getByRole('button', { name: 'PDF file', exact: true }).click();
      const result = await Promise.race([pending, page.locator('.error-banner').waitFor({ state: 'visible' }).then(async () => {
        throw Error(await page.locator('.error-banner').innerText());
      })]);
      await result.saveAs(root + '/direct.pdf'); outputs['direct.pdf'] = hash(await fs.readFile(root + '/direct.pdf'));
    } catch (error) { failure = String(error); }
    expect(await model()).toEqual(before);
    await fs.writeFile(root + '/browser-report.json', JSON.stringify({ outputs, errors, failure, paint,
      sourceHash: hash(input), authoringReceiptHash: hash(await fs.readFile(inputRoot + '/browser-report.json')),
      buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-script-scope-print.spec.ts')),
      referenceHash: hash(await fs.readFile('tests/fixtures/native-word-script-scopes.json')) }, null, 2));
    expect(errors).toEqual([]); expect(failure).toBeNull();
  });
