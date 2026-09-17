import { expect, test } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { paintPresentationFixture } from './fixtures/pptx-paints';
test('PowerPoint imports and edits text fills and transparency with inherited theme/background preservation', async ({
  page,
}) => {
  const input = Buffer.from(await paintPresentationFixture());
  await fs.mkdir('.local/pptx-validation', { recursive: true });
  await fs.writeFile('.local/pptx-validation/paints-source.pptx', input);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Paint editing.pptx',
    mimeType: 'application/octet-stream',
    buffer: input,
  });
  const stage = page.locator('.slide-stage');
  const text = stage.getByRole('group', { name: 'Filled text', exact: true });
  await expect(stage).toHaveCSS('background-color', 'rgb(187, 221, 238)');
  await expect(text).toHaveCSS('background-color', 'rgba(221, 238, 170, 0.5)');
  await expect(stage.locator('.slide-element').nth(1)).toHaveCSS(
    'background-color',
    'rgba(0, 0, 0, 0)',
  );
  await expect(stage.locator('.slide-element').nth(2)).toHaveCSS(
    'background-color',
    'rgb(52, 86, 120)',
  );
  await text.click();
  await page.getByLabel('Shape fill', { exact: true }).fill('#223344');
  await page.getByRole('spinbutton', { name: 'Fill transparency', exact: true }).fill('25');
  await page.getByRole('spinbutton', { name: 'Fill transparency', exact: true }).press('Tab');
  await expect(text).toHaveCSS('background-color', 'rgba(34, 51, 68, 0.75)');
  await page.getByRole('button', { name: 'No fill', exact: true }).click();
  await expect(text).toHaveCSS('background-color', 'rgba(0, 0, 0, 0)');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(text).toHaveCSS('background-color', 'rgba(34, 51, 68, 0.75)');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Paint editing', exact: true }).click();
  await expect(text).toHaveCSS('background-color', 'rgba(34, 51, 68, 0.75)');
  await page.getByRole('button', { name: 'Present', exact: true }).click();
  await expect(
    page.locator('.presentation-mode foreignObject > div').filter({ hasText: 'Filled text' }),
  ).toHaveCSS('background-color', 'rgba(34, 51, 68, 0.75)');
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.writeFile('.local/pptx-validation/paints-edited.pptx', output);
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Paint result.pptx',
    mimeType: 'application/octet-stream',
    buffer: output,
  });
  await expect(text).toHaveCSS('background-color', 'rgba(34, 51, 68, 0.75)');
  await expect(stage).toHaveCSS('background-color', 'rgb(187, 221, 238)');
  await page.screenshot({ path: 'test-results/pptx-paints.png' });
});
