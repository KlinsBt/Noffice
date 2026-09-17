import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordPlacements, clickWordText } from './word-pagination-helpers';

for (const name of ['mixed', 'single', 'keep'] as const)
  test(`Word variable pagination: ${name} editing, history, reload and actual exports`, async ({
    page,
  }) => {
    test.setTimeout(120000);
    const root = `.local/word-variable-${name}`;
    await fs.mkdir(root, { recursive: true });
    const source = await fs.readFile(`tests/fixtures/word-variable-${name}.docx`);
    const native = JSON.parse(
      await fs.readFile('tests/fixtures/native-word-variable-pagination.json', 'utf8'),
    );
    const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
    expect(hash(source)).toBe(native.cases[name].sourceSha256);
    const expected = (key: string) =>
      native.cases[key].snapshot.paragraphs.map((p: any) =>
        p.lines.flatMap((line: any) =>
          [...line.text.replaceAll('\v', '\n')].map((text) => ({ text, page: line.page })),
        ),
      );
    const upload = async (label: string, bytes: Buffer) => {
      await page.locator('input[type=file][multiple]').setInputFiles({
        name: `${label}.docx`,
        buffer: bytes,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
        label,
      );
    };
    await page.setViewportSize({ width: 1500, height: 1200 });
    await page.goto('/');
    await upload(`Variable ${name}`, source);
    const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
    const paragraphs = editor.locator('p');
    const clickWord = (word: string) => clickWordText(page, editor, word);
    const check = async (key = name as string) => {
      await expect(paragraphs).toHaveCount(12);
      await expect(page.locator('.section-page')).toHaveCount(native.cases[key].snapshot.pages);
      await expect.poll(() => wordPlacements(editor)).toEqual(expected(key));
    };
    const download = async (stage: string) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download');
      await page
        .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
        .click();
      const bytes = await fs.readFile((await (await pending).path())!);
      await fs.writeFile(`${root}/${stage}.docx`, bytes);
      return bytes;
    };
    await check();
    const positions = await wordPlacements(editor);
    // A pointer on a later fragment edits the semantic paragraph, including at 50%.
    for (let i = 0; i < 5; i++)
      await page.getByRole('button', { name: 'Zoom out', exact: true }).click();
    await check();
    await clickWord('p3c');
    await page.keyboard.press('End');
    await page.keyboard.insertText('EDIT');
    await expect(paragraphs.nth(2)).toContainText('EDIT');
    const changed = await wordPlacements(editor);
    expect(changed[2].filter((c) => 'EDIT'.includes(c.text)).every((c) => c.page === 2)).toBe(true);
    await page.keyboard.press('Control+z');
    await check();
    for (let i = 0; i < 5; i++)
      await page.getByRole('button', { name: 'Zoom in', exact: true }).click();
    await check();
    // Invalid spacing must not enter history or disturb the page allocation.
    await paragraphs.first().click();
    await page.getByRole('combobox', { name: 'Line spacing', exact: true }).selectOption('custom');
    const dialog = page.getByRole('dialog', { name: 'Line spacing', exact: true });
    await dialog.getByRole('spinbutton', { name: 'Spacing amount', exact: true }).fill('0');
    await expect(dialog.getByRole('button', { name: 'Apply', exact: true })).toBeDisabled();
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
    await check();
    let keepHash: string | undefined, keepPositions: unknown;
    if (name === 'mixed') {
      const commands = [
        [1, 'Keep with next'],
        [2, 'Keep lines together'],
        [3, 'Page break before'],
        [8, 'Widow/orphan control'],
        [9, 'Keep with next'],
      ] as const;
      for (const [index, label] of commands) {
        await clickWord(`p${index + 1}a`);
        await page.getByRole('button', { name: 'Layout', exact: true }).click();
        await page.getByRole('checkbox', { name: label, exact: true }).check();
      }
      await check('keep');
      keepPositions = await wordPlacements(editor);
      for (let i = 0; i < 5; i++)
        await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await check();
      for (let i = 0; i < 5; i++)
        await page.getByRole('button', { name: 'Redo', exact: true }).click();
      await check('keep');
      const output = await download('browser-keep');
      keepHash = hash(output);
      await page.pdf({ path: `${root}/browser-keep-print.pdf`, preferCSSPageSize: true });
      await upload('Variable keep reimport', output);
      await check('keep');
      await upload(`Variable ${name}`, source);
      await check();
      await page.getByRole('button', { name: 'Home', exact: true }).click();
    }
    await clickWord('p12d');
    await page.keyboard.press('Control+End');
    await page.keyboard.insertText('!');
    await expect(paragraphs.last()).toContainText('p12d!');
    const editedPositions = await wordPlacements(editor);
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await check();
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page
      .getByRole('button', { name: `Variable ${name}`, exact: true })
      .first()
      .click();
    await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
    const edited = await download('browser-edited');
    await page.pdf({ path: `${root}/browser-edited-print.pdf`, preferCSSPageSize: true });
    await upload(`Variable ${name} reimport`, edited);
    await expect.poll(() => wordPlacements(editor)).toEqual(editedPositions);
    await clickWord('p12d');
    await page.keyboard.press('Control+End');
    await page.keyboard.press('Backspace');
    await check();
    const restored = await download('browser');
    await page.pdf({ path: `${root}/browser-print.pdf`, preferCSSPageSize: true });
    await upload(`Variable ${name} restored`, restored);
    await check();
    const stored = await page.evaluate(async () => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const r = indexedDB.open('noffice-workspace');
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
      try {
        return JSON.stringify(
          await new Promise((resolve, reject) => {
            const r = db.transaction('files').objectStore('files').getAll();
            r.onsuccess = () => resolve(r.result);
            r.onerror = () => reject(r.error);
          }),
        );
      } finally {
        db.close();
      }
    });
    expect(stored).not.toMatch(
      /word-page-fragment|word-line-from|word-baseline-shift|word-softline-strut/,
    );
    await page.screenshot({ path: `${root}/screen.png` });
    await fs.writeFile(
      `${root}/browser-report.json`,
      JSON.stringify(
        {
          passed: true,
          sourceHash: hash(source),
          exportHash: hash(restored),
          editedHash: hash(edited),
          positions,
          editedPositions,
          keepHash,
          keepPositions,
          browser: page.context().browser()!.version(),
        },
        null,
        2,
      ) + '\n',
    );
  });
