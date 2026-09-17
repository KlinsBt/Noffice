import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

test('Word font-size boundaries survive import, editing, history, reload and downloaded exports', async ({
  page,
}) => {
  test.setTimeout(90000);
  const root = '.local/word-font-sizes';
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/word-font-sizes.docx');
  await fs.writeFile(`${root}/source.docx`, source);
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const upload = async (buffer: Buffer, name: string) => {
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: `${name}.docx`,
      buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  await page.goto('/');
  await upload(source, 'Font boundaries');
  const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
  const sizes = [1, 1.5, 6, 7.5, 10, 160, 160.5, 200, 1638];
  const readSizes = () =>
    editor
      .locator('p')
      .evaluateAll((ps) =>
        ps.map((p) => parseFloat(getComputedStyle(p.querySelector('span') || p).fontSize) * 0.75),
      );
  await expect(editor.locator('p')).toHaveCount(sizes.length);
  (await readSizes()).forEach((size, i) => expect(size).toBeCloseTo(sizes[i], 2));
  const input = page.getByRole('spinbutton', { name: 'Font size', exact: true });
  const selectTarget = async () => {
    await editor.focus();
    await page.keyboard.press('Control+Home');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Shift+ArrowRight');
  };
  const exports: { file: string; sha256: string; size: number }[] = [];
  for (const [button, size] of [
    ['Grow font', 2.5],
    ['Shrink font', 1],
  ] as const) {
    await selectTarget();
    await page.getByRole('button', { name: button, exact: true }).click();
    expect((await readSizes())[1]).toBeCloseTo(size, 2);
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    expect((await readSizes())[1]).toBeCloseTo(1.5, 2);
  }
  let previous = 1.5;
  for (const size of [6, 200, 1.5, 1638]) {
    await selectTarget();
    await input.fill(String(size));
    await input.press('Tab');
    expect((await readSizes())[1]).toBeCloseTo(size, 2);
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    expect((await readSizes())[1]).toBeCloseTo(previous, 2);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    expect((await readSizes())[1]).toBeCloseTo(size, 2);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    const name = await page.getByRole('textbox', { name: 'File name', exact: true }).inputValue();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    expect((await readSizes())[1]).toBeCloseTo(size, 2);
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    const file = `browser-${size}.docx`;
    await fs.writeFile(`${root}/${file}`, bytes);
    exports.push({ file, sha256: hash(bytes), size });
    await upload(bytes, `Font boundaries ${size}`);
    const actual = await readSizes();
    actual.forEach((value, i) => expect(value).toBeCloseTo(i === 1 ? size : sizes[i], 2));
    previous = size;
  }
  for (const invalid of ['0', '1638.5', '1.25', '']) {
    await selectTarget();
    await input.fill(invalid);
    await input.press('Tab');
    expect((await readSizes())[1]).toBeCloseTo(previous, 2);
    await expect(input).toHaveValue(String(previous));
  }
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceHash: hash(source),
        exports,
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ),
  );
});
