import { expect, test } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { retainedPresentationFixture, pictureData } from './fixtures/pptx-retained';
import { flippedImageFixture } from './fixtures/pptx-retained';
test('imported image flips render in canvas and slideshow and survive undo, reload and PPTX export', async ({
  page,
}) => {
  const input = await flippedImageFixture();
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Image flips.pptx',
    mimeType: 'application/octet-stream',
    buffer: Buffer.from(input),
  });
  const object = page.getByRole('group', { name: 'image object', exact: true });
  await object.click();
  await expect(object.locator('img')).toHaveCSS('transform', 'matrix(-1, 0, 0, 1, 0, 0)');
  await page.getByRole('button', { name: 'Flip horizontally', exact: true }).click();
  await page.getByRole('button', { name: 'Flip vertically', exact: true }).click();
  await expect(object.locator('img')).toHaveCSS('transform', 'matrix(1, 0, 0, -1, 0, 0)');
  await page.keyboard.press('Control+z');
  await expect(object.locator('img')).toHaveCSS('transform', 'matrix(1, 0, 0, 1, 0, 0)');
  await page.keyboard.press('Control+y');
  await expect(object.locator('img')).toHaveCSS('transform', 'matrix(1, 0, 0, -1, 0, 0)');
  await page.getByRole('spinbutton', { name: 'Object rotation', exact: true }).fill('30');
  await page.getByLabel('Speaker notes', { exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Image flips', exact: true }).click();
  await expect(object.locator('img')).toHaveCSS('transform', 'matrix(1, 0, 0, -1, 0, 0)');
  await page.getByRole('button', { name: 'Present', exact: true }).click();
  await expect(
    page.getByRole('dialog', { name: 'Slideshow', exact: true }).locator('image'),
  ).toHaveAttribute('transform', /scale\(1 -1\)/);
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.mkdir('.local/pptx-validation', { recursive: true });
  await fs.writeFile('.local/pptx-validation/flipped.pptx', output);
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Flipped result.pptx',
    mimeType: 'application/octet-stream',
    buffer: output,
  });
  await expect(object.locator('img')).toHaveCSS('transform', 'matrix(1, 0, 0, -1, 0, 0)');
});
test('imported slide reordering survives reload and preserves slide and chart relationships', async ({
  page,
}) => {
  const input = await retainedPresentationFixture();
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Reordered.pptx',
    mimeType: 'application/octet-stream',
    buffer: Buffer.from(input),
  });
  await page.getByRole('group', { name: 'Original title', exact: true }).waitFor();
  await page.getByRole('button', { name: 'Move slide down', exact: true }).click();
  await page.getByRole('button', { name: 'Slide 1', exact: true }).click();
  await expect(page.getByLabel('Speaker notes', { exact: true })).toHaveValue('Chart notes');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Reordered', exact: true }).click();
  await expect(page.getByLabel('Speaker notes', { exact: true })).toHaveValue('Chart notes');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.mkdir('.local/pptx-validation', { recursive: true });
  await fs.writeFile('.local/pptx-validation/reordered.pptx', output);
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/presentation.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Reordered result.pptx',
    mimeType: 'application/octet-stream',
    buffer: output,
  });
  await expect(page.getByLabel('Speaker notes', { exact: true })).toHaveValue('Chart notes');
  await page.getByRole('button', { name: 'Slide 2', exact: true }).click();
  await expect(page.getByRole('group', { name: 'Original title', exact: true })).toBeVisible();
});
test('existing PowerPoint edits retain source parts through save/reload and download', async ({
  page,
}) => {
  const input = await retainedPresentationFixture();
  await page.goto('/');
  await page.getByRole('button', { name: 'Start PowerPoint', exact: true }).waitFor();
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Source.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: Buffer.from(input),
  });
  await expect(page.getByText('Standard · 4:3', { exact: true })).toBeVisible();
  const canvas = await page.getByRole('application', { name: 'Slide canvas' }).boundingBox();
  expect(canvas!.width / canvas!.height).toBeCloseTo(4 / 3, 3);
  await expect(page.locator('.slide-thumb svg').first()).toHaveAttribute('viewBox', '0 0 960 720');
  await page.getByRole('group', { name: 'Original title', exact: true }).dblclick();
  await page
    .getByRole('textbox', { name: 'Edit slide text', exact: true })
    .fill('Browser revised title');
  await page.getByLabel('Speaker notes', { exact: true }).click();
  await page.getByRole('group', { name: 'Browser revised title', exact: true }).click();
  await page.getByRole('spinbutton', { name: 'Object x', exact: true }).fill('150');
  await page.getByRole('spinbutton', { name: 'Object rotation', exact: true }).fill('30');
  await page.getByRole('spinbutton', { name: 'Object rotation', exact: true }).press('Tab');
  await expect(page.getByRole('group', { name: 'Browser revised title', exact: true })).toHaveCSS(
    'transform',
    /matrix\(/,
  );
  await page.getByRole('application', { name: 'Slide canvas' }).click({ position: { x: 2, y: 2 } });
  await page.keyboard.press('Control+z');
  await expect(page.getByRole('group', { name: 'Browser revised title', exact: true })).toHaveCSS(
    'transform',
    'none',
  );
  await page.keyboard.press('Control+y');
  await expect(page.getByRole('group', { name: 'Browser revised title', exact: true })).toHaveCSS(
    'transform',
    /matrix\(/,
  );
  await page.getByRole('group', { name: 'Browser revised title', exact: true }).click();
  await page.getByRole('combobox', { name: 'Font family', exact: true }).selectOption('Inter');
  await page.getByRole('button', { name: 'Underline text', exact: true }).click();
  await page.getByLabel('Speaker notes', { exact: true }).fill('Browser revised notes');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Source', exact: true }).click();
  await expect(
    page.getByRole('group', { name: 'Browser revised title', exact: true }),
  ).toBeVisible();
  await expect(page.getByText('Standard · 4:3', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Present', exact: true }).click();
  await expect(
    page.getByRole('dialog', { name: 'Slideshow' }).locator('svg.slide-preview > g').first(),
  ).toHaveAttribute('transform', /^rotate\(30 /);
  await expect(
    page.getByRole('dialog', { name: 'Slideshow' }).locator('svg.slide-preview'),
  ).toHaveAttribute('viewBox', '0 0 960 720');
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.mkdir('.local/pptx-validation', { recursive: true });
  await fs.writeFile('.local/pptx-validation/edited.pptx', output);
  const after = await JSZip.loadAsync(output),
    before = await JSZip.loadAsync(input);
  for (const path of Object.keys(before.files))
    if (
      !before.files[path].dir &&
      !['ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml'].includes(path)
    )
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await page.screenshot({ path: 'test-results/pptx-retained-edit.png' });
});
test('an imported deck accepts a local image and its first speaker notes without rebuilding', async ({
  page,
}) => {
  const input = await retainedPresentationFixture(false);
  await page.goto('/');
  await page.getByRole('button', { name: 'Start PowerPoint', exact: true }).waitFor();
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'No notes.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: Buffer.from(input),
  });
  await page.getByLabel('Speaker notes', { exact: true }).fill('First browser notes');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByLabel('Speaker notes', { exact: true })).toHaveValue('');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(page.getByLabel('Speaker notes', { exact: true })).toHaveValue(
    'First browser notes',
  );
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'No notes', exact: true }).click();
  await expect(page.getByLabel('Speaker notes', { exact: true })).toHaveValue(
    'First browser notes',
  );
  await page.locator('input[type=file][accept^="image/png"]').setInputFiles({
    name: 'Pixel.png',
    mimeType: 'image/png',
    buffer: Buffer.from(pictureData, 'base64'),
  });
  await expect(page.locator('.slide-element.image')).toHaveCount(1);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.mkdir('.local/pptx-validation', { recursive: true });
  await fs.writeFile('.local/pptx-validation/additions.pptx', output);
  const after = await JSZip.loadAsync(output),
    before = await JSZip.loadAsync(input);
  for (const path of Object.keys(before.files))
    if (
      !before.files[path].dir &&
      ![
        'ppt/slides/slide1.xml',
        'ppt/slides/_rels/slide1.xml.rels',
        'ppt/presentation.xml',
        'ppt/_rels/presentation.xml.rels',
        '[Content_Types].xml',
      ].includes(path)
    )
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  expect(
    Object.keys(after.files).filter((p) => /^ppt\/notesMasters\/[^/]+\.xml$/.test(p)),
  ).toHaveLength(1);
  expect(await after.file('ppt/notesSlides/_rels/notesSlide1.xml.rels')!.async('string')).toContain(
    '/notesMaster',
  );
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles({
      name: 'Notes result.pptx',
      mimeType: 'application/octet-stream',
      buffer: output,
    });
  await expect(page.getByLabel('Speaker notes', { exact: true })).toHaveValue(
    'First browser notes',
  );
  await expect(page.locator('.slide-element.image')).toHaveCount(1);
});
