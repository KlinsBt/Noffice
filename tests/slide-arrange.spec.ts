import { expect, test, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';

async function fixture() {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'ARRANGE', width: 10, height: 7.5 });
  pptx.layout = 'ARRANGE';
  const slide = pptx.addSlide();
  [
    [60, 40, 180, 60, 30],
    [310, 200, 70, 120, 90],
    [570, 390, 130, 90, 270],
  ].forEach(([x, y, w, h, rotate], i) =>
    slide.addText(['First object', 'Second object', 'Third object'][i], {
      x: x / 72,
      y: y / 72,
      w: w / 72,
      h: h / 72,
      rotate,
      fontFace: 'Arial',
      fontSize: 18,
      color: '234567',
      margin: 0,
    }),
  );
  slide.addNotes('Arrangement preserves notes');
  pptx.addSlide().addText('Untouched slide', { x: 1, y: 1, w: 5, h: 1 });
  return Buffer.from((await pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer);
}
const objects = (page: Page) => page.locator('.slide-stage > .slide-element');
async function positions(page: Page) {
  return objects(page).evaluateAll((nodes) =>
    nodes.map((node) => ({
      x: parseFloat((node as HTMLElement).style.left),
      y: parseFloat((node as HTMLElement).style.top),
    })),
  );
}
async function assertPositions(page: Page, points: number[][]) {
  const actual = await positions(page);
  points.forEach(([x, y], i) => {
    expect(actual[i].x).toBeCloseTo((x * 4) / 3, 2);
    expect(actual[i].y).toBeCloseTo((y * 4) / 3, 2);
  });
}

test('PowerPoint multi-selection alignment, distribution and movement retain native geometry through saved PPTX export', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1440, height: 1000 });
  const input = await fixture();
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Arrange objects.pptx',
    mimeType: 'application/octet-stream',
    buffer: input,
  });
  await objects(page).nth(0).click();
  await objects(page)
    .nth(1)
    .click({ modifiers: ['Control'] });
  await objects(page)
    .nth(2)
    .click({ modifiers: ['Shift'] });
  await expect(page.locator('.slide-stage > .selected')).toHaveCount(3);
  await expect(page.getByText('3 objects selected', { exact: true })).toBeVisible();
  const arrange = page.getByRole('combobox', { name: 'Arrange objects', exact: true });
  await arrange.selectOption('left');
  await assertPositions(page, [
    [60, 40],
    [82.05764, 200],
    [37.05764, 390],
  ]);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await assertPositions(page, [
    [60, 40],
    [310, 200],
    [570, 390],
  ]);
  await arrange.selectOption('horizontal');
  await assertPositions(page, [
    [60, 40],
    [380, 200],
    [570, 390],
  ]);
  await page
    .getByRole('combobox', { name: 'Align relative to', exact: true })
    .selectOption('slide');
  await arrange.selectOption('horizontal');
  await assertPositions(page, [
    [82.5, 40],
    [370, 200],
    [527.5, 390],
  ]);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.keyboard.press('Control+y');
  await assertPositions(page, [
    [82.5, 40],
    [370, 200],
    [527.5, 390],
  ]);

  // A drag previews and commits the whole selection in one undo step.
  const beforeDrag = await positions(page),
    box = (await objects(page).nth(0).boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 + 24, box.y + box.height / 2 + 12, { steps: 3 });
  await page.mouse.up();
  const afterDrag = await positions(page);
  const dx = afterDrag[0].x - beforeDrag[0].x,
    dy = afterDrag[0].y - beforeDrag[0].y;
  expect(dx).toBeGreaterThan(10);
  expect(dy).toBeGreaterThan(5);
  afterDrag.forEach((pos, i) => {
    // Browsers serialize CSS coordinates to a limited number of significant digits.
    expect(pos.x - beforeDrag[i].x).toBeCloseTo(dx, 2);
    expect(pos.y - beforeDrag[i].y).toBeCloseTo(dy, 2);
  });
  await page.keyboard.press('Control+z');
  expect(await positions(page)).toEqual(beforeDrag);
  await page.keyboard.press('Escape');
  await expect(page.locator('.slide-stage > .selected')).toHaveCount(0);
  await page.keyboard.press('Control+a');
  await expect(page.locator('.slide-stage > .selected')).toHaveCount(3);
  await arrange.selectOption('center');
  await arrange.selectOption('bottom');
  await page.getByRole('button', { name: 'Selection pane', exact: true }).click();
  await expect(page.getByRole('button', { name: '1. First object', exact: true })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  await page.getByRole('button', { name: '1. First object', exact: true }).click();
  await expect(page.locator('.slide-stage > .selected')).toHaveCount(2);
  await page.getByRole('button', { name: '1. First object', exact: true }).click();
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('ArrowDown');
  await assertPositions(page, [
    [277.5, 440.0192],
    [332.5, 446],
    [302.5, 431],
  ]);
  await page.keyboard.press('Delete');
  await expect(
    page.getByText(
      'Deleting imported objects requires relationship-aware editing and is not supported yet.',
      { exact: true },
    ),
  ).toBeVisible();
  await expect(objects(page)).toHaveCount(3);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.screenshot({ path: 'test-results/powerpoint-arrange-desktop.png' });
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Arrange objects', exact: true }).click();
  await assertPositions(page, [
    [277.5, 440.0192],
    [332.5, 446],
    [302.5, 431],
  ]);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.mkdir('.local/pptx-validation', { recursive: true });
  await fs.writeFile('.local/pptx-validation/arranged.pptx', output);
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Arranged result.pptx',
    mimeType: 'application/octet-stream',
    buffer: output,
  });
  await assertPositions(page, [
    [277.5, 440.0192],
    [332.5, 446],
    [302.5, 431],
  ]);
});

test('selection pane supports keyboard toggles, slide changes and narrow windows', async ({
  page,
}) => {
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Selection.pptx',
    mimeType: 'application/octet-stream',
    buffer: await fixture(),
  });
  await page.getByRole('button', { name: 'Selection pane', exact: true }).click();
  await page.getByRole('button', { name: '1. First object', exact: true }).focus();
  await page.keyboard.press('Space');
  await page.keyboard.press('Tab');
  await page.keyboard.press('Space');
  await expect(page.locator('.slide-stage > .selected')).toHaveCount(2);
  await page.getByRole('button', { name: 'Slide 2', exact: true }).click();
  await expect(page.locator('.slide-stage > .selected')).toHaveCount(0);
  await expect(page.getByRole('button', { name: '1. Untouched slide', exact: true })).toBeVisible();
  await page.setViewportSize({ width: 390, height: 844 });
  await page.getByRole('button', { name: '1. Untouched slide', exact: true }).click();
  await expect(page.locator('.slide-stage > .selected')).toHaveCount(1);
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1)).toBe(
    true,
  );
});
