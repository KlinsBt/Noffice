import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
const native: typeof import('./fixtures/native-powerpoint-child-bounds.json') = JSON.parse(
  await fs.readFile('tests/fixtures/native-powerpoint-child-bounds.json', 'utf8'),
);
const ui: typeof import('./fixtures/native-powerpoint-child-ui.json') = JSON.parse(
  await fs.readFile('tests/fixtures/native-powerpoint-child-ui.json', 'utf8'),
);

test('rotated sibling bounds and native size-field coupling survive history, recovery and actual exports', async ({
  page,
}) => {
  test.setTimeout(180000);
  const root = '.local/powerpoint-child-bounds';
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile('tests/fixtures/powerpoint-child-bounds.pptx');
  const sha = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  expect(sha(source)).toBe(native.sourceSha256);
  const button = (name: string) => page.getByRole('button', { name, exact: true });
  const objects = page.locator('.slide-stage [role=group]'),
    sibling = objects.nth(1);
  const field = (key: string) =>
    page.getByRole('spinbutton', { name: `Object ${key}`, exact: true });
  const upload = async (buffer: Buffer, name: string) => {
    await page
      .locator('input[type=file][multiple]')
      .setInputFiles({ name: `${name}.pptx`, mimeType: 'application/octet-stream', buffer });
    await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(name);
  };
  const check = async (frame: { x: number; y: number; w: number; h: number }) => {
    const actual = await sibling.evaluate((el) => {
      const s = getComputedStyle(el);
      return {
        x: parseFloat(s.left),
        y: parseFloat(s.top),
        w: parseFloat(s.width),
        h: parseFloat(s.height),
      };
    });
    for (const key of ['x', 'y', 'w', 'h'] as const)
      expect(Math.abs(actual[key] - (frame[key] * 4) / 3), key).toBeLessThan(1 / 64 + 0.001);
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
  await upload(source, 'Child bounds edits');
  await sibling.click();
  await field('w').fill('');
  await expect(page.getByText('Enter a finite size or position.', { exact: true })).toBeVisible();
  await check(native.initial[0].find((s) => s.id === 3)!);
  await field('w').fill('120');
  await objects.first().click();
  await field('rotation').fill('45');
  await field('rotation').press('Tab');
  await expect(
    page.getByText(
      'This edit requires unsupported grouped text, skew, rotation or reflection changes.',
      { exact: true },
    ),
  ).toBeVisible();
  for (let i = 0; i < 4; i++) {
    await button(`Slide ${i + 1}`).click();
    await sibling.click();
    await check(native.initial[i].find((s) => s.id === 3)!);
    const stages = native.steps.filter((s) => s.slide === i + 1);
    for (const [j, key] of (['x', 'y', 'w', 'h'] as const).entries()) {
      const expected = stages[j].shapes.find((s) => s.id === 3)!;
      await field(key).fill(String((expected[key] * 4) / 3));
      await check(expected);
    }
    for (let n = 0; n < 4; n++) await button('Undo').click();
    await check(native.initial[i].find((s) => s.id === 3)!);
    for (let n = 0; n < 4; n++) await button('Redo').click();
    await check(native.effective[i].find((s) => s.id === 3)!);
  }
  await page.getByLabel('Speaker notes', { exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const saved = await page.evaluate(async () => {
    const open = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => {
      open.onsuccess = () => resolve(open.result);
    });
    const get = db.transaction('files').objectStore('files').getAll();
    const files = await new Promise<any[]>((resolve) => {
      get.onsuccess = () => resolve(get.result);
    });
    const file = files.find((f) => f.name === 'Child bounds edits');
    for (const slide of file.content.slides)
      for (const el of slide.elements)
        if (el.sourceGroupTransform) delete el.sourceGroupTransform.dimensions;
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
  await button('Child bounds edits').click();
  for (let i = 0; i < 4; i++) {
    await button(`Slide ${i + 1}`).click();
    await check(native.effective[i].find((s) => s.id === 3)!);
  }
  const revision = await page.evaluate(async (id) => {
    const r = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => {
      r.onsuccess = () => resolve(r.result);
    });
    const g = db.transaction('files').objectStore('files').get(id);
    const f = await new Promise<any>((resolve) => {
      g.onsuccess = () => resolve(g.result);
    });
    db.close();
    return f.revision;
  }, saved.id);
  expect(revision).toBe(saved.revision);
  await button('Slide 3').click();
  await page.screenshot({ path: `${root}/browser-edits.png` });
  await button('Present').click();
  await expect(page.getByRole('dialog', { name: 'Slideshow', exact: true })).toBeVisible();
  await page.keyboard.press('Escape');
  const edits = await output('edits');
  await upload(edits, 'Child bounds reimport');
  for (let i = 0; i < 4; i++) {
    await button(`Slide ${i + 1}`).click();
    await check(native.effective[i].find((s) => s.id === 3)!);
  }
  await upload(source, 'Child bounds UI');
  await button('Slide 3').click();
  await sibling.click();
  await check(ui.stages[0].frame);
  for (const [i, key] of (['x', 'y', 'w', 'h'] as const).entries()) {
    await field(key).fill(String((ui.stages[i + 1].frame[key] * 4) / 3));
    await check(ui.stages[i + 1].frame);
  }
  for (let i = 5; i <= 12; i++) {
    await button(ui.stages[i].stage).click();
    await check(ui.stages[i].frame);
  }
  await page.getByLabel('Speaker notes', { exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await button('Recent files').click();
  await button('Child bounds UI').click();
  await button('Slide 3').click();
  await check(ui.stages[13].frame);
  const uiOutput = await output('ui');
  await upload(uiOutput, 'Child UI reimport');
  await button('Slide 3').click();
  await check(ui.stages[13].frame);
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceSha256: sha(source),
        editsSha256: sha(edits),
        uiSha256: sha(uiOutput),
        browser: page.context().browser()!.version(),
      },
      null,
      2,
    ) + '\n',
  );
});
