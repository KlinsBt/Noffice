import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

export function storyKeyboardTests(
  modes: readonly string[],
  prefix: string,
  output: string,
  options: { kinds?: readonly ('header' | 'footer')[]; text?: string } = {},
) {
  for (const mode of modes)
    for (const kind of options.kinds || ['header', 'footer'])
      test(`Word ${mode} ${kind} keyboard typing fully undoes and redoes before export`, async ({
        page,
      }) => {
        const name = `${mode}-${kind}`;
        const root = `${output}/${name}`;
        await fs.mkdir(root, { recursive: true });
        const source = await fs.readFile(`tests/fixtures/${prefix}-${mode}.docx`);
        const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
        const errors: string[] = [];
        page.on('pageerror', (e) => errors.push(e.message));
        await page.goto('/');
        await page.locator('input[type=file][multiple]').setInputFiles({
          name: `${name}.docx`,
          buffer: source,
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });
        await expect(page.locator('.section-page')).toHaveCount(6);
        const open = async () => {
          await page.getByRole('button', { name: 'Insert', exact: true }).click();
          await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
          await page
            .getByRole('button', { name: new RegExp(`^Section 1 — First-page ${kind}`) })
            .click();
          return page.getByRole('textbox', { name: 'Header or footer text', exact: true });
        };
        const story = await open();
        const text = options.text || `Created ${kind}s 2`;
        await story.focus();
        await page.keyboard.type(text);
        await expect(story).toHaveText(text);
        const typed = await story.textContent();
        await page.keyboard.press('Control+z');
        await expect(story).not.toHaveText(text);
        const undone = await story.textContent();
        await page.keyboard.press('Control+y');
        await expect(story).toHaveText(text);
        const undoAll: string[] = [],
          redoAll: string[] = [];
        while (await story.textContent()) {
          expect(undoAll.length).toBeLessThan(32);
          await page.keyboard.press('Control+z');
          undoAll.push((await story.textContent())!);
        }
        for (let i = 0; i < undoAll.length; i++) {
          await page.keyboard.press('Control+y');
          redoAll.push((await story.textContent())!);
        }
        await expect(story).toHaveText(text);
        await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
        await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
        await page.reload();
        await page.getByRole('button', { name: 'Recent files', exact: true }).click();
        await page.getByRole('button', { name, exact: true }).click();
        await expect(await open()).toHaveText(text);
        await page.getByRole('button', { name: 'Cancel', exact: true }).click();
        await page.getByRole('button', { name: 'Export', exact: true }).click();
        const download = page.waitForEvent('download');
        await page
          .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
          .click();
        await (await download).saveAs(`${root}/keyboard.docx`);
        expect(errors).toEqual([]);
        await fs.writeFile(
          `${root}/browser-report.json`,
          JSON.stringify(
            {
              sourceHash: hash(source),
              exportHash: hash(await fs.readFile(`${root}/keyboard.docx`)),
              buildHash: await wordStoryBuildHash(),
              typed,
              undone,
              undoAll,
              redoAll,
              errors,
            },
            null,
            2,
          ),
        );
      });
}
