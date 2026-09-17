import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

const native = JSON.parse(
  await fs.readFile('tests/fixtures/native-powerpoint-coarse-ui.json', 'utf8'),
);
const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');

test('coarse group commits match native position history, recover failures and export editable frames', async ({
  page,
}) => {
  const root = '.local/powerpoint-coarse';
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/powerpoint-coarse.pptx');
  expect(hash(source)).toBe(native.sourceSha256);
  const button = (name: string) => page.getByRole('button', { name, exact: true });
  const object = page.locator('.slide-stage [role=group]').first();
  const field = (key: string) =>
    page.getByRole('spinbutton', { name: `Object ${key}`, exact: true });
  const upload = async (buffer: Buffer, name: string) => {
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({ name: `${name}.pptx`, mimeType: 'application/octet-stream', buffer });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  const check = async (frame: Record<string, number>) => {
    const actual = await object.evaluate((el) => {
      const s = getComputedStyle(el);
      return {
        x: parseFloat(s.left),
        y: parseFloat(s.top),
        w: parseFloat(s.width),
        h: parseFloat(s.height),
      };
    });
    for (const key of ['x', 'y', 'w', 'h'])
      expect(Math.abs(actual[key as keyof typeof actual] - (frame[key] * 4) / 3)).toBeLessThan(
        1 / 64 + 0.001,
      );
  };
  const commit = async (key: string, value: string) => {
    await field(key).fill(value);
    await field(key).press('Tab');
  };
  const output = async (stage: string) => {
    await button('Export').click();
    const pending = page.waitForEvent('download');
    await button('PPTX file Editable in Microsoft PowerPoint').click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${stage}.pptx`, bytes);
    return bytes;
  };
  await page.goto('/');
  await upload(source, 'Coarse groups');
  await object.click();
  await check(native.stages[0].frame);
  await commit('x', '161');
  await check(native.stages[0].frame);
  await expect(field('x')).toHaveValue('160');
  await expect(button('Undo')).toBeDisabled();
  await object.click();
  await page.keyboard.press('ArrowRight');
  await check(native.stages[0].frame);
  await expect(button('Undo')).toBeDisabled();
  const bounds = (await object.boundingBox())!;
  const start = { x: bounds.x + bounds.width / 2, y: bounds.y + bounds.height / 2 };
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(start.x + (5 * bounds.width) / 160, start.y, { steps: 3 });
  await page.mouse.up();
  await check(native.stages[1].frame);
  await button('Undo').click();
  await check(native.stages[0].frame);
  await commit('x', '');
  await expect(page.getByText('Enter a finite size or position.', { exact: true })).toBeVisible();
  await check(native.stages[0].frame);
  await expect(button('Undo')).toBeDisabled();
  await field('x').fill('');
  await field('x').pressSequentially(String((121.6 * 4) / 3));
  await check(native.stages[0].frame);
  await field('x').press('Tab');
  await check(native.stages[1].frame);
  await commit('y', String((141.25 * 4) / 3));
  await check(native.stages[2].frame);
  for (const i of [3, 4, 5, 6]) {
    await button(native.stages[i].stage).click();
    await check(native.stages[i].frame);
  }
  await page.getByLabel('Speaker notes', { exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  // Seed an edited pre-grid IndexedDB record. Migration must retain edits and revision.
  const saved = await page.evaluate(async () => {
    const request = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => {
      request.onsuccess = () => resolve(request.result);
    });
    const read = db.transaction('files').objectStore('files').getAll();
    const files = await new Promise<any[]>((resolve) => {
      read.onsuccess = () => resolve(read.result);
    });
    const file = files.find((f) => f.name === 'Coarse groups');
    for (const slide of file.content.slides)
      for (const el of slide.elements) delete el.sourceGroupTransform.dimensions;
    const tx = db.transaction('files', 'readwrite');
    tx.objectStore('files').put(file);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
    return { id: file.id, revision: file.revision };
  });
  await page.reload();
  await button('Recent files').click();
  await button('Coarse groups').click();
  await check(native.stages[7].frame);
  const position = await output('position');
  const revision = await page.evaluate(async (id) => {
    const r = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => {
      r.onsuccess = () => resolve(r.result);
    });
    const read = db.transaction('files').objectStore('files').get(id);
    const record = await new Promise<any>((resolve) => {
      read.onsuccess = () => resolve(read.result);
    });
    db.close();
    return record.revision;
  }, saved.id);
  expect(revision).toBe(saved.revision);
  await object.click();
  await commit('w', String((121.6 * 4) / 3));
  await commit('h', String((68.75 * 4) / 3));
  const resized = { ...native.stages[7].frame, w: 123, h: 69.75 };
  await check(resized);
  await button('Undo').click();
  await button('Undo').click();
  await check(native.stages[7].frame);
  await button('Redo').click();
  await button('Redo').click();
  await check(resized);
  for (const i of [3, 4]) {
    await button(`Slide ${i}`).click();
    await object.click();
    await commit('x', String(Number(await field('x').inputValue()) + 16));
    await expect(
      page.getByText(
        'Coarse coordinates in rotated or reflected groups cannot be edited accurately yet.',
        { exact: true },
      ),
    ).toBeVisible();
  }
  await button('Slide 1').click();
  await check(resized);
  await page.screenshot({ path: `${root}/browser-size.png` });
  await button('Present').click();
  await expect(
    page.getByRole('dialog', { name: 'Slideshow', exact: true }).locator('rect[width="164"]'),
  ).toBeVisible();
  await page.keyboard.press('Escape');
  const size = await output('size');
  for (const [name, bytes, frame] of [
    ['Position result', position, native.stages[7].frame],
    ['Size result', size, resized],
  ] as const) {
    await upload(bytes, name);
    await check(frame);
  }
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceSha256: hash(source),
        positionSha256: hash(position),
        sizeSha256: hash(size),
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ),
  );
});
