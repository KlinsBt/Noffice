import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
const native: typeof import('./fixtures/native-powerpoint-group-edits.json') = JSON.parse(
  await fs.readFile('tests/fixtures/native-powerpoint-group-edits.json', 'utf8'),
);
const original: typeof import('./fixtures/native-powerpoint-groups.json') = JSON.parse(
  await fs.readFile('tests/fixtures/native-powerpoint-groups.json', 'utf8'),
);

test('rotated and reflected group children move/resize through history, reload and actual exports', async ({
  page,
}) => {
  const root = '.local/powerpoint-group-edits';
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/powerpoint-groups.pptx');
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  expect(hash(source)).toBe(native.sourceSha256);
  const button = (name: string) => page.getByRole('button', { name, exact: true });
  const objects = page.locator('.slide-stage [role=group]');
  const upload = async (buffer: Buffer, name: string) => {
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({ name: `${name}.pptx`, mimeType: 'application/octet-stream', buffer });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  const check = async (expected: (typeof native.effective)[number]) => {
    for (const [j, item] of expected.entries()) {
      const frame = await objects.nth(j).evaluate((el) => {
        const s = getComputedStyle(el);
        return {
          x: parseFloat(s.left),
          y: parseFloat(s.top),
          w: parseFloat(s.width),
          h: parseFloat(s.height),
        };
      });
      for (const key of ['x', 'y', 'w', 'h'] as const)
        expect(Math.abs(frame[key] - (item[key] * 4) / 3)).toBeLessThan(1 / 64 + 0.001);
    }
  };
  await page.goto('/');
  await upload(source, 'Transformed children');
  for (const i of [2, 3]) {
    await button(`Slide ${i + 1}`).click();
    await check(original.effective[i]);
    await objects.first().click();
    // Final-frame workflow; native size-field coupling is covered separately.
    for (const key of ['w', 'h', 'x', 'y'] as const)
      await page
        .getByRole('spinbutton', { name: `Object ${key}`, exact: true })
        .fill(String((native.effective[i][0][key] * 4) / 3));
    await check(native.effective[i]);
    for (let n = 0; n < 4; n++) await button('Undo').click();
    await check(original.effective[i]);
    for (let n = 0; n < 4; n++) await button('Redo').click();
    await check(native.effective[i]);
  }
  await page.getByLabel('Speaker notes', { exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await button('Recent files').click();
  await button('Transformed children').click();
  for (const i of [2, 3]) {
    await button(`Slide ${i + 1}`).click();
    await check(native.effective[i]);
  }
  await button('Present').click();
  const show = page.getByRole('dialog', { name: 'Slideshow', exact: true });
  await expect(show.locator('rect[width="176"]')).toBeVisible();
  await page.keyboard.press('Escape');
  await button('Export').click();
  const pending = page.waitForEvent('download');
  await button('PPTX file Editable in Microsoft PowerPoint').click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.writeFile(`${root}/edited.pptx`, output);
  await upload(output, 'Transformed result');
  for (let i = 0; i < 4; i++) {
    await button(`Slide ${i + 1}`).click();
    await check(native.effective[i]);
  }
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceSha256: hash(source),
        exportSha256: hash(output),
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ),
  );
});
