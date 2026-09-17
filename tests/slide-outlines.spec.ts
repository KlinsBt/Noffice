import { expect, test } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { paintPresentationFixture } from './fixtures/pptx-paints';
test('PowerPoint outlines render and edit with undo, persistence and retained export', async ({
  page,
}) => {
  const input = Buffer.from(await paintPresentationFixture());
  await fs.mkdir('.local/pptx-validation', { recursive: true });
  await fs.writeFile('.local/pptx-validation/outlines-source.pptx', input);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Outline editing.pptx',
    mimeType: 'application/octet-stream',
    buffer: input,
  });
  const shape = page.locator('.slide-stage .slide-element').nth(1),
    outline = shape.locator('.object-outline');
  await expect(outline).toHaveAttribute('stroke', '#ff0000');
  await expect(outline).toHaveAttribute('stroke-width', '2');
  await shape.click();
  await page.getByLabel('Outline color', { exact: true }).fill('#224466');
  await page.getByLabel('Outline width', { exact: true }).fill('6');
  await page.getByLabel('Outline width', { exact: true }).press('Tab');
  await page.getByLabel('Outline transparency', { exact: true }).fill('40');
  await page.getByLabel('Outline transparency', { exact: true }).press('Tab');
  await page.getByLabel('Outline dashes', { exact: true }).selectOption('dashDot');
  await expect(outline).toHaveAttribute('stroke', 'rgba(34, 68, 102, 0.6)');
  await expect(outline).toHaveAttribute('stroke-dasharray', '24 18 6 18');
  await page.getByRole('button', { name: 'No outline', exact: true }).click();
  await expect(outline).toHaveCount(0);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(outline).toHaveAttribute('stroke-width', '6');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Outline editing', exact: true }).click();
  await expect(outline).toHaveAttribute('stroke-dasharray', '24 18 6 18');
  await page.getByRole('button', { name: 'Present', exact: true }).click();
  await expect(
    page.locator('.presentation-mode .object-outline[stroke-width="6"]'),
  ).toHaveAttribute('stroke', 'rgba(34, 68, 102, 0.6)');
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.writeFile('.local/pptx-validation/outlines-edited.pptx', output);
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Outline result.pptx',
    mimeType: 'application/octet-stream',
    buffer: output,
  });
  await expect(outline).toHaveAttribute('stroke-width', '6');
  await expect(outline).toHaveAttribute('stroke', 'rgba(34, 68, 102, 0.6)');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await shape.click();
  await page.screenshot({ path: 'test-results/pptx-outlines.png' });
});
