import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
const native = JSON.parse(
  await fs.readFile('tests/fixtures/native-powerpoint-groups.json', 'utf8'),
);

const root = '.local/powerpoint-groups';
const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
test('native nested group frames, child editing, history, reload, export and rejection', async ({
  page,
}) => {
  const source = await fs.readFile('tests/fixtures/powerpoint-groups.pptx');
  await fs.mkdir(root, { recursive: true });
  const upload = (buffer: Buffer, name: string) =>
    page
      .locator('input[type=file][multiple]')
      .setInputFiles({ name: `${name}.pptx`, mimeType: 'application/octet-stream', buffer });
  const download = async (name: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${name}`, bytes);
    return bytes;
  };
  await page.goto('/');
  await upload(source, 'Native groups');
  const objects = page.locator('.slide-stage [role=group]');
  for (let i = 0; i < 4; i++) {
    await page.getByRole('button', { name: `Slide ${i + 1}`, exact: true }).click();
    for (let j = 0; j < 3; j++) {
      const frame = await objects.nth(j).evaluate((el) => {
        const s = getComputedStyle(el);
        const matrix = new DOMMatrix(s.transform);
        return {
          x: parseFloat(s.left),
          y: parseFloat(s.top),
          w: parseFloat(s.width),
          h: parseFloat(s.height),
          rotation: (Math.atan2(matrix.b, matrix.a) * 180) / Math.PI,
        };
      });
      for (const key of ['x', 'y', 'w', 'h'] as const)
        // Chromium lays out dimensions on a 1/64 CSS-pixel grid; the importer has tighter unit assertions.
        expect(Math.abs(frame[key] - (native.effective[i][j][key] * 4) / 3)).toBeLessThan(
          1 / 64 + 0.001,
        );
      expect(frame.rotation).toBeCloseTo(native.effective[i][j].rotation, 3);
    }
  }
  expect(await download('unchanged.pptx')).toEqual(source);
  // Moving a reflected child is supported; changing its rotation remains guarded.
  await objects.first().click();
  const x = page.getByRole('spinbutton', { name: 'Object x', exact: true });
  await page.getByRole('spinbutton', { name: 'Object rotation', exact: true }).fill('45');
  await page.keyboard.press('Tab');
  await expect(
    page.getByText(
      'This edit requires unsupported grouped text, skew, rotation or reflection changes.',
      { exact: true },
    ),
  ).toBeVisible();
  await expect(objects.first()).toHaveCSS('left', '533.333px');
  await page.getByRole('button', { name: 'Slide 1', exact: true }).click();
  await objects.first().click();
  await x.fill('176');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(objects.first()).toHaveCSS('left', '160px');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(objects.first()).toHaveCSS('left', '176px');
  for (const [name, value] of [
    ['Object y', '198.66666666666666'],
    ['Object w', '176'],
    ['Object h', '102'],
  ])
    await page.getByRole('spinbutton', { name, exact: true }).fill(value);
  await page.getByLabel('Speaker notes', { exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Native groups', exact: true }).click();
  await expect(objects.first()).toHaveCSS('left', '176px');
  await expect(objects.first()).toHaveCSS('height', '102px');
  await page.getByRole('button', { name: 'Present', exact: true }).click();
  await expect(
    page.getByRole('dialog', { name: 'Slideshow', exact: true }).locator('rect[x="176"]'),
  ).toBeVisible();
  await page.keyboard.press('Escape');
  const edited = await download('edited.pptx');
  await upload(edited, 'Group result');
  await expect(objects.first()).toHaveCSS('left', '176px');
  await expect(objects.first()).toHaveCSS('width', '176px');
  await expect(objects.first()).toHaveCSS('height', '102px');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.screenshot({ path: `${root}/browser.png` });
  await fs.writeFile(
    `${root}/browser-report.json`,
    JSON.stringify(
      {
        passed: true,
        sourceSha256: hash(source),
        exportSha256: hash(edited),
        browser: page.context().browser()!.version(),
        checks:
          '12 native leaf frames, nested child move/resize, undo/redo, reload, slideshow, unchanged/edited export, reimport and rejected grouped-child rotation edit',
      },
      null,
      2,
    ),
  );
});

test('legacy IndexedDB group frames migrate without changing source bytes or the storage revision', async ({
  page,
}) => {
  const source = await fs.readFile('tests/fixtures/powerpoint-groups.pptx');
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Legacy groups.pptx',
    mimeType: 'application/octet-stream',
    buffer: source,
  });
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const saved = await page.evaluate(async () => {
    const request = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => {
      request.onsuccess = () => resolve(request.result);
    });
    const read = db.transaction('files').objectStore('files').getAll();
    const files = await new Promise<any[]>((resolve) => {
      read.onsuccess = () => resolve(read.result);
    });
    const file = files.find((f) => f.name === 'Legacy groups');
    delete file.content.groupModelVersion;
    for (const slide of file.content.slides)
      for (const el of slide.elements) {
        const local = el.sourceGroupTransform.local;
        Object.assign(el, {
          x: (local.x / 9144000) * 960,
          y: (local.y / 5143500) * 540,
          w: (local.w / 9144000) * 960,
          h: (local.h / 5143500) * 540,
          rotation: local.rotation,
        });
        delete el.sourceGroupTransform;
      }
    const canonical = (v: any): any =>
      Array.isArray(v)
        ? v.map(canonical)
        : v && typeof v === 'object'
          ? Object.fromEntries(
              Object.entries(v)
                .filter(([, x]) => x !== undefined)
                .sort(([a], [b]) => a.localeCompare(b))
                .map(([k, x]) => [k, canonical(x)]),
            )
          : v;
    const bytes = new TextEncoder().encode(JSON.stringify(canonical(file.content)));
    file.original.contentFingerprint = Array.from(
      new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)),
      (b) => b.toString(16).padStart(2, '0'),
    ).join('');
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
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Legacy groups', exact: true }).click();
  await expect(page.locator('.slide-stage [role=group]').first()).toHaveCSS('left', '160px');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
    .click();
  expect(await fs.readFile((await (await pending).path())!)).toEqual(source);
  const revision = await page.evaluate(async (id) => {
    const r = indexedDB.open('noffice-workspace', 1);
    const db = await new Promise<IDBDatabase>((resolve) => {
      r.onsuccess = () => resolve(r.result);
    });
    const request = db.transaction('files').objectStore('files').get(id);
    const record = await new Promise<any>((resolve) => {
      request.onsuccess = () => resolve(request.result);
    });
    db.close();
    return record.revision;
  }, saved.id);
  expect(revision).toBe(saved.revision);
});
