import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { waitWordLayout, wordPlacements } from './word-pagination-helpers';

// Every case owns its browser storage and output directory. Use the existing
// four-worker limit for this independent matrix instead of serializing 77 cases.
test.describe.configure({ mode: 'parallel' });

for (const length of ['short', 'long', 'boundary'] as const) {
  const matrix =
    length === 'long'
      ? 'word-inter-long-metrics'
      : length === 'boundary'
        ? 'word-leading-boundary'
        : 'word-inter-metrics';
  const native = JSON.parse(readFileSync(`tests/fixtures/native-${matrix}.json`, 'utf8')) as {
    cases: {
      name: string;
      sourceHash: string;
      starts: { page: number; y?: number; text?: string }[];
    }[];
  };

  for (const sample of native.cases)
    test(`Word regular Inter native line matrix ${length}: ${sample.name}`, async ({ page }) => {
      const finalText = sample.starts.at(-1)?.text ?? 'sample12';
      const root = `.local/${matrix}/${sample.name}`;
      await fs.mkdir(root, { recursive: true });
      const source = await fs.readFile(`tests/fixtures/${matrix}/${sample.name}.docx`);
      expect(createHash('sha256').update(source).digest('hex')).toBe(sample.sourceHash);
      await page.goto('/');
      await page.locator('input[type=file][multiple]').setInputFiles({
        name: `${sample.name}.docx`,
        buffer: source,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
      await expect(editor).toContainText(finalText);
      await waitWordLayout(editor);
      const initial = await wordPlacements(editor);
      const linePages = initial.flatMap((paragraph) =>
        paragraph.filter((c, i, all) => i === 0 || all[i - 1].text === '\n').map((c) => c.page),
      );
      expect(linePages).toEqual(sample.starts.map((c) => c.page));
      await editor.focus();
      await page.keyboard.press('Control+End');
      await page.keyboard.insertText('!');
      await expect(editor).toContainText(`${finalText}!`);
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await expect(editor).not.toContainText('!');
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await expect(editor).toContainText(`${finalText}!`);
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      await page.reload();
      await page.getByRole('button', { name: 'Recent files', exact: true }).click();
      await page.getByRole('button', { name: sample.name, exact: true }).click();
      await waitWordLayout(editor);
      expect(await wordPlacements(editor)).toEqual(initial);
      const outputs: Record<string, string> = {};
      for (const kind of ['docx', 'pdf'] as const) {
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        const pending = page.waitForEvent('download');
        await page
          .getByRole('button', {
            name: kind === 'pdf' ? 'PDF file' : 'DOCX file Editable in Microsoft Word',
            exact: true,
          })
          .click();
        const bytes = await fs.readFile((await (await pending).path())!);
        await fs.writeFile(`${root}/restored.${kind}`, bytes);
        outputs[kind] = createHash('sha256').update(bytes).digest('hex');
      }
      await fs.writeFile(
        `${root}/browser-report.json`,
        JSON.stringify(
          {
            passed: true,
            sample: sample.name,
            sourceHash: sample.sourceHash,
            initial,
            linePages,
            outputs,
            browser: page.context().browser()!.version(),
          },
          null,
          2,
        ) + '\n',
      );
    });
}
