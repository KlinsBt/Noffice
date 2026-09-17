import { expect, test } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { stackPresentationFixture } from './fixtures/pptx-stack';

test('PowerPoint layering retains interleaved source objects through multiselection, undo, reload and actual PPTX export', async ({
  page,
}) => {
  const input = Buffer.from(await stackPresentationFixture());
  await page.setViewportSize({ width: 1440, height: 1000 });
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Layered deck.pptx',
    mimeType: 'application/octet-stream',
    buffer: input,
  });
  const objects = page.locator('.slide-stage > .slide-element');
  const order = () =>
    objects.evaluateAll((nodes) => nodes.map((node) => node.getAttribute('aria-label')));
  await expect.poll(order).toEqual(['Bottom text', 'image object', 'Top text']);
  await page.getByRole('button', { name: 'Selection pane', exact: true }).click();
  const pane = page.locator('.object-selection');
  await pane.getByRole('button', { name: /Bottom text$/ }).click();
  await pane.getByRole('button', { name: /Top text$/ }).click();
  const command = page.getByRole('combobox', { name: 'Order objects', exact: true });
  await command.selectOption('front');
  await expect.poll(order).toEqual(['image object', 'Bottom text', 'Top text']);
  await expect(page.locator('.slide-stage > .selected')).toHaveCount(2);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect.poll(order).toEqual(['Bottom text', 'image object', 'Top text']);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect.poll(order).toEqual(['image object', 'Bottom text', 'Top text']);
  await command.selectOption('back');
  await expect.poll(order).toEqual(['Bottom text', 'Top text', 'image object']);
  await command.selectOption('forward');
  await expect.poll(order).toEqual(['image object', 'Bottom text', 'Top text']);
  await command.selectOption('backward');
  await expect.poll(order).toEqual(['Bottom text', 'Top text', 'image object']);
  await command.selectOption('front');
  await pane.getByRole('button', { name: /Top text$/ }).click();
  await page.getByRole('button', { name: 'Bring to front', exact: true }).click();
  await page.getByRole('textbox', { name: 'Object text', exact: true }).fill('Bottom text edited');
  await page.getByRole('button', { name: 'Text box', exact: true }).click();
  await page.getByRole('textbox', { name: 'Object text', exact: true }).fill('Inserted layer');
  await command.selectOption('back');
  await expect
    .poll(order)
    .toEqual(['Inserted layer', 'image object', 'Top text', 'Bottom text edited']);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Layered deck', exact: true }).click();
  await expect
    .poll(order)
    .toEqual(['Inserted layer', 'image object', 'Top text', 'Bottom text edited']);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.mkdir('.local/pptx-validation', { recursive: true });
  await fs.writeFile('.local/pptx-validation/stacked.pptx', output);
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Layer result.pptx',
    mimeType: 'application/octet-stream',
    buffer: output,
  });
  await expect
    .poll(order)
    .toEqual(['Inserted layer', 'image object', 'Top text', 'Bottom text edited']);
  await expect(page.getByLabel('Speaker notes', { exact: true })).toHaveValue(
    'Layering retains notes',
  );
  await page.screenshot({ path: 'test-results/pptx-stack.png' });
});
