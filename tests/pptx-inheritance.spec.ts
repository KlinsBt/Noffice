import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';

const root = '.local/powerpoint-inheritance';
const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
test('native inherited and rich placeholders retain styles through edits, history, reload and actual export', async ({
  page,
}) => {
  const source = await fs.readFile('tests/fixtures/powerpoint-inheritance.pptx');
  await fs.mkdir(root, { recursive: true });
  const upload = async (buffer: Buffer, name: string) =>
    page
      .locator('input[type=file][multiple]')
      .setInputFiles({ name: `${name}.pptx`, mimeType: 'application/octet-stream', buffer });
  const download = async (file: string) => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
      .click();
    const bytes = await fs.readFile((await (await pending).path())!);
    await fs.writeFile(`${root}/${file}`, bytes);
    return bytes;
  };
  await page.goto('/');
  await upload(source, 'Native inheritance');
  const stage = page.locator('.slide-stage');
  const title = stage.getByRole('group', { name: 'Master 1 title', exact: true });
  await expect(title.locator('span')).toHaveCSS('font-size', '50.6667px');
  await expect(title.locator('span')).toHaveCSS('font-style', 'italic');
  await expect(title.locator('span')).toHaveCSS(
    'font-family',
    '"Calibri Light", Inter, sans-serif',
  );
  await expect(title.locator('.slide-text-paragraph')).toHaveCSS('text-align', 'center');
  await page.getByRole('button', { name: 'Slide 2', exact: true }).click();
  await expect(
    stage.getByRole('group', { name: 'Master 2 title', exact: true }).locator('span'),
  ).toHaveCSS('font-size', '53.3333px');
  await page.getByRole('button', { name: 'Slide 1', exact: true }).click();
  expect(await download('unchanged.pptx')).toEqual(source);
  await title.click();
  await page.getByRole('spinbutton', { name: 'Object x', exact: true }).fill('96');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(title).toHaveCSS('left', '80px');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(title).toHaveCSS('left', '96px');
  await page
    .getByRole('textbox', { name: 'Object text', exact: true })
    .fill('Master 1 title edited');
  let rich = stage.getByRole('group', { name: 'Plain BOLD tail', exact: true });
  await expect(rich.locator('span').nth(1)).toHaveCSS('font-weight', '700');
  await expect(rich.locator('span').nth(1)).toHaveCSS('font-size', '37.3333px');
  await expect(rich).toHaveCSS('text-decoration-line', 'none');
  await expect(rich.locator('span').nth(0)).toHaveCSS('text-decoration-line', 'underline');
  await expect(rich.locator('span').nth(1)).toHaveCSS('text-decoration-line', 'none');
  await expect(rich.locator('span').nth(2)).toHaveCSS('font-style', 'italic');
  await expect(rich.locator('span').nth(2)).toHaveCSS('color', 'rgb(17, 34, 51)');
  await rich.click();
  await page.getByRole('textbox', { name: 'Object text', exact: true }).fill('Plain BRAVE tail');
  rich = stage.getByRole('group', { name: 'Plain BRAVE tail', exact: true });
  await expect(rich.locator('span').nth(1)).toHaveText('BRAVE');
  await expect(rich.locator('span').nth(1)).toHaveCSS('font-weight', '700');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(stage.getByRole('group', { name: 'Plain BOLD tail', exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(rich).toBeVisible();
  await page.getByLabel('Speaker notes', { exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Native inheritance', exact: true }).click();
  await expect(rich.locator('span').nth(1)).toHaveCSS('font-size', '37.3333px');
  await page.getByRole('button', { name: 'Present', exact: true }).click();
  const show = page.getByRole('dialog', { name: 'Slideshow', exact: true });
  await expect(show.locator('span').filter({ hasText: /^BRAVE$/ })).toHaveCSS('font-weight', '700');
  await page.keyboard.press('Escape');
  const edited = await download('edited.pptx');
  const before = await JSZip.loadAsync(source),
    after = await JSZip.loadAsync(edited);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await upload(edited, 'Inheritance result');
  await expect(rich.locator('span').nth(1)).toHaveCSS('font-weight', '700');
  await expect(stage.getByRole('group', { name: 'Master 1 title edited', exact: true })).toHaveCSS(
    'left',
    '96px',
  );
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
          'Two layouts, inherited fonts, rich runs, text/geometry edits, undo/redo, reload, slideshow, unchanged bytes and unrelated package payloads',
      },
      null,
      2,
    ),
  );
});

test('legacy saved decks hydrate inherited fonts and rich runs without a storage revision or changed export', async ({
  page,
}) => {
  const source = await fs.readFile('tests/fixtures/powerpoint-inheritance.pptx');
  await page.goto('/');
  await page
    .locator('input[type=file][multiple]')
    .setInputFiles({
      name: 'Legacy inheritance.pptx',
      mimeType: 'application/octet-stream',
      buffer: source,
    });
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const previous = await page.evaluate(async () => {
    const db = await new Promise<IDBDatabase>((resolve) => {
      const r = indexedDB.open('noffice-workspace', 1);
      r.onsuccess = () => resolve(r.result);
    });
    const files = await new Promise<any[]>((resolve) => {
      const r = db.transaction('files').objectStore('files').getAll();
      r.onsuccess = () => resolve(r.result);
    });
    const file = files.find((f) => f.name === 'Legacy inheritance');
    delete file.content.textModelVersion;
    for (const slide of file.content.slides) {
      delete slide.sourceLayoutPath;
      delete slide.sourceMasterPath;
      delete slide.sourceThemePath;
      for (const element of slide.elements) {
        delete element.sourceText;
        delete element.sourcePlaceholder;
        delete element.sourceGroupIds;
        element.align = 'left';
        if (element.text.startsWith('Master'))
          Object.assign(element, {
            fontSize: 32,
            fontFamily: undefined,
            italic: false,
            color: '#263d34',
          });
      }
    }
    const canonical = (value: any): any =>
      Array.isArray(value)
        ? value.map(canonical)
        : value && typeof value === 'object'
          ? Object.fromEntries(
              Object.entries(value)
                .filter(([, v]) => v !== undefined)
                .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
                .map(([k, v]) => [k, canonical(v)]),
            )
          : value;
    const digest = await crypto.subtle.digest(
      'SHA-256',
      new TextEncoder().encode(JSON.stringify(canonical(file.content))),
    );
    file.original.contentFingerprint = [...new Uint8Array(digest)]
      .map((v) => v.toString(16).padStart(2, '0'))
      .join('');
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction('files', 'readwrite');
      tx.objectStore('files').put(file);
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
    return { id: file.id, revision: file.revision };
  });
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Legacy inheritance', exact: true }).click();
  const title = page
    .locator('.slide-stage')
    .getByRole('group', { name: 'Master 1 title', exact: true });
  await expect(title.locator('span')).toHaveCSS('font-size', '50.6667px');
  await expect(
    page
      .locator('.slide-stage')
      .getByRole('group', { name: 'Plain BOLD tail', exact: true })
      .locator('span')
      .nth(1),
  ).toHaveCSS('font-weight', '700');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
    .click();
  expect(await fs.readFile((await (await pending).path())!)).toEqual(source);
  const revision = await page.evaluate(async (id) => {
    const db = await new Promise<IDBDatabase>((resolve) => {
      const r = indexedDB.open('noffice-workspace', 1);
      r.onsuccess = () => resolve(r.result);
    });
    const value = await new Promise<any>((resolve) => {
      const r = db.transaction('files').objectStore('files').get(id);
      r.onsuccess = () => resolve(r.result);
    });
    db.close();
    return value.revision;
  }, previous.id);
  expect(revision).toBe(previous.revision);
});
