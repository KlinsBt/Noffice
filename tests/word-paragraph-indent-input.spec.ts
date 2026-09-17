import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-paragraph-indent-input.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
const fields: Record<string, string> = { LeftIndent: 'Indent before text', RightIndent: 'Indent after text', FirstLineIndent: 'First line (negative for hanging)' };
const cases = [{ property: 'LeftIndent', input: .775 }, { property: 'FirstLineIndent', input: -.625 },
  { property: 'RightIndent', input: 720.025 }, { property: 'LeftIndent', input: 1584.00001 },
  { property: 'RightIndent', input: -1584 }, { property: 'FirstLineIndent', input: -1584 }];
for (const context of ['body', 'header']) for (const sample of cases)
  test(`Word ${context} indent ${sample.property} ${sample.input}: boundaries, history, reload and DOCX`, async ({ page }) => {
    const native = reference.rows.find(r => r.context === context && r.property === sample.property && r.input === sample.input)!;
    expect(native?.accepted).toBe(true);
    const run = process.env.NOFFICE_PARAGRAPH_INDENT_RUN || 'indent-browser-v1'; expect(run).toMatch(/^indent-browser-v\d+$/);
    const root = `.local/word-story-paragraph-layout/${run}/${native.name}`; await fs.mkdir(root, { recursive: true });
    const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
    const source = await fs.readFile('tests/fixtures/word-story-paragraph-layout-header-indent-start-source.docx');
    expect(hash(source)).toBe(reference.originalHash);
    const name = 'Indent input ' + native.name, errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
    await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const scope = context === 'header' ? page.getByRole('dialog') : page;
    const open = async () => {
      if (context === 'header') {
        await page.getByRole('button', { name: 'Insert', exact: true }).click();
        await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
        await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
        const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
        await editor.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.press('Control+ArrowDown'); await page.keyboard.press('Control+ArrowDown');
      } else {
        await page.getByRole('button', { name: 'Home', exact: true }).click();
        await page.getByRole('textbox', { name: 'Document text', exact: true }).focus(); await page.keyboard.press('Control+Home');
        await page.getByRole('button', { name: 'Layout', exact: true }).click();
      }
      return scope.getByRole('spinbutton', { name: fields[sample.property], exact: true });
    };
    const history = async (label: 'Undo' | 'Redo', empty = false) => {
      if (context === 'body') await page.getByRole('button', { name: 'Home', exact: true }).click();
      await scope.getByRole('button', { name: label, exact: true }).click();
      if (empty) await expect(scope.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
      if (context === 'body') await page.getByRole('button', { name: 'Layout', exact: true }).click();
    };
    const control = await open(); await expect(control).toHaveValue('0');
    for (const invalid of ['', '1585', '-1585', '100000']) {
      await control.fill(invalid); await control.press('Enter'); await expect(control).toHaveAttribute('aria-invalid', 'true');
      await control.press('Escape'); await expect(control).toHaveValue('0');
    }
    await control.fill(String(sample.input)); await control.press('Enter');
    const wanted = String(native.savedTwips / 20); await expect(control).toHaveValue(wanted);
    await control.fill(wanted); await control.press('Enter');
    await history('Undo', true); await expect(control).toHaveValue('0'); await history('Redo'); await expect(control).toHaveValue(wanted);
    if (context === 'header') await scope.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    const outputs: Record<string, string> = {};
    const download = async (stage: string) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click(); const waiting = page.waitForEvent('download');
      await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();
      await (await waiting).saveAs(root + '/' + stage + '.docx'); outputs[stage + '.docx'] = hash(await fs.readFile(root + '/' + stage + '.docx'));
    };
    await download('edited'); await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload(); await page.getByRole('button', { name: 'Recent files', exact: true }).click(); await page.getByRole('button', { name, exact: true }).click();
    await expect(await open()).toHaveValue(wanted);
    if (context === 'header') await scope.getByRole('button', { name: 'Cancel', exact: true }).click();
    await download('reloaded'); await page.goto('/'); await page.locator('input[type=file][multiple]').setInputFiles(root + '/edited.docx');
    await expect(await open()).toHaveValue(wanted);
    if (context === 'header') await scope.getByRole('button', { name: 'Cancel', exact: true }).click();
    expect(errors).toEqual([]);
    await fs.writeFile(root + '/browser-report.json', JSON.stringify({ name: native.name, context, sample, wanted, outputs, sourceHash: hash(source), errors,
      buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-paragraph-indent-input.spec.ts')), referenceHash: hash(await fs.readFile('tests/fixtures/native-word-paragraph-indent-input.json')) }, null, 2));
  });
