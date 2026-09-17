import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import native from './fixtures/native-word-paragraph-fonts.json' with { type: 'json' };

for (const sample of native.cases)
  test(`Word paragraph font ${sample.name}: commands, history, reload and actual DOCX`, async ({
    page,
  }) => {
    const root = `.local/word-paragraph-fonts/${sample.name}`;
    await fs.mkdir(root, { recursive: true });
    await page.goto('/');
    await page.getByRole('button', { name: 'Start Word', exact: true }).click();
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    await editor.focus();
    for (const [i, paragraph] of sample.before.paragraphs.entries()) {
      if (i) await page.keyboard.press('Enter');
      const text = paragraph.text.replace(/\r$/, '');
      if (text) await page.keyboard.insertText(text);
    }
    await page.keyboard.press('Control+a');
    const family = page.getByRole('combobox', { name: 'Font family', exact: true });
    const size = page.getByRole('spinbutton', { name: 'Font size', exact: true });
    await family.selectOption('Arial');
    await size.fill('13');
    await size.press('Tab');
    await size.fill('12');
    await size.press('Tab');
    const snapshot = () =>
      editor.locator('p').evaluateAll((ps) =>
        ps.map((p) => {
          // CSS serializes 20pt as 26.6667px; retain 0.0001pt precision.
          const points = (value: string) =>
            Number((parseFloat(value) * (value.endsWith('px') ? 0.75 : 1)).toFixed(4));
          const font = (value: string) => value.replace(/^(["'])(.*)\1$/, '$2');
          const mark = {
            font: font((p as HTMLElement).style.fontFamily),
            size: points((p as HTMLElement).style.fontSize),
          };
          const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
          const characters: { text: string; font: string; size: number }[] = [];
          for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            const style = getComputedStyle(node.parentElement!);
            for (const text of node.textContent!)
              characters.push({ text, font: font(style.fontFamily), size: points(style.fontSize) });
          }
          return { text: p.textContent + '\r', mark, characters };
        }),
      );
    const baseline = await snapshot();
    const outputs: Record<string, string> = {};
    const download = async (stage: string) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page
        .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
        .click();
      const bytes = await fs.readFile((await (await pending).path())!);
      await fs.writeFile(`${root}/${stage}.docx`, bytes);
      outputs[stage] = createHash('sha256').update(bytes).digest('hex');
    };
    await download('source');
    await editor.focus();
    await page.keyboard.press('Control+Home');
    for (let i = 0; i < sample.selection.start; i++) await page.keyboard.press('ArrowRight');
    for (let i = sample.selection.start; i < sample.selection.end; i++)
      await page.keyboard.press('Shift+ArrowRight');
    await family.selectOption('Courier New');
    await size.fill('20');
    await size.press('Tab');
    await expect.poll(snapshot).toEqual(sample.after.paragraphs);
    await download('edited');
    if (sample.name !== 'caret-at-start') {
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await expect.poll(snapshot).toEqual(baseline);
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await expect.poll(snapshot).toEqual(sample.after.paragraphs);
    }
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name: 'Untitled document', exact: true }).click();
    await expect.poll(snapshot).toEqual(sample.after.paragraphs);
    await download('restored');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: 'Font reimport.docx',
      buffer: await fs.readFile(`${root}/restored.docx`),
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
      'Font reimport',
    );
    await expect.poll(snapshot).toEqual(sample.after.paragraphs);
    const typing = native.typingCases.find((c) => c.name === sample.name);
    if (typing) {
      await editor.focus();
      await page.keyboard.press('Control+Home');
      for (let i = 0; i < sample.selection.end; i++) await page.keyboard.press('ArrowRight');
      await page.keyboard.insertText('!');
      await expect.poll(snapshot).toEqual(typing.after.paragraphs);
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await expect.poll(snapshot).toEqual(sample.after.paragraphs);
      await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await expect.poll(snapshot).toEqual(typing.after.paragraphs);
      await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
      await page.reload();
      await page.getByRole('button', { name: 'Recent files', exact: true }).click();
      await page.getByRole('button', { name: 'Font reimport', exact: true }).click();
      await expect.poll(snapshot).toEqual(typing.after.paragraphs);
      await download('typed');
    }
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify(
        {
          passed: true,
          outputs,
          selection: sample.selection,
          expected: sample.after,
          browser: page.context().browser()!.version(),
        },
        null,
        2,
      ),
    );
  });
