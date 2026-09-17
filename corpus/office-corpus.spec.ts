import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import manifest from './manifest.json' with { type: 'json' };

const root = path.resolve('.local/office-corpus');
const marker = 'Noffice corpus edit';
const application = { word: 'Word', excel: 'Excel', powerpoint: 'PowerPoint' };
async function exported(page: Page, kind: string, ext: string) {
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const download = page.waitForEvent('download', { timeout: 15000 });
  const error = page.getByRole('alert').filter({ hasText: 'Export failed:' });
  const outcome = Promise.race([
    download,
    error.waitFor({ timeout: 15000 }).then(async () => {
      throw new Error((await error.textContent()) || 'Export failed');
    }),
  ]);
  await page
    .getByRole('button', {
      name: `${ext.toUpperCase()} file Editable in Microsoft ${application[kind as keyof typeof application]}`,
      exact: true,
    })
    .click();
  return fs.readFile((await (await outcome).path())!);
}
for (const fixture of manifest.files) {
  test(`${fixture.kind}: ${fixture.filename} — ${fixture.focus}`, async ({ page }) => {
    const directory = path.join(root, 'results', fixture.id);
    await fs.mkdir(directory, { recursive: true });
    const suffix = fixture.filename.split('.').pop()!;
    for (const stage of ['unchanged', 'edited'])
      await fs.rm(path.join(directory, `${stage}.${suffix}`), { force: true });
    const result: Record<string, unknown> = {
      id: fixture.id,
      kind: fixture.kind,
      source: fixture.source,
      stages: [],
      browserErrors: [],
      externalRequests: [],
    };
    const stages = result.stages as string[],
      errors = result.browserErrors as string[],
      external = result.externalRequests as string[];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('**/*', async (route) => {
      const url = new URL(route.request().url());
      if (['http:', 'https:'].includes(url.protocol) && url.hostname !== '127.0.0.1') {
        external.push(url.origin);
        await route.abort();
      } else await route.continue();
    });
    try {
      const sourcePath = path.join(root, 'input', fixture.kind, fixture.filename);
      const source = await fs.readFile(sourcePath);
      expect(
        createHash('sha1')
          .update(Buffer.concat([Buffer.from(`blob ${source.length}\0`), source]))
          .digest('hex'),
      ).toBe(fixture.gitBlobSha1);
      const ext = fixture.filename.split('.').pop()!;
      await page.goto('/');
      await page.locator('input[type=file][multiple]').setInputFiles(sourcePath);
      if (fixture.id === 'poi-docx-deep-table-cell') {
        await expect(page.getByRole('alert')).toContainText(
          'Document XML nesting exceeds the supported 256 levels.',
        );
        expect(errors).toEqual([]);
        expect(external).toEqual([]);
        stages.push('bounded-rejection');
        result.status = 'safely-rejected';
        return;
      }
      await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
        fixture.filename.slice(0, -(ext.length + 1)),
        { timeout: 20000 },
      );
      stages.push('import');
      await page.screenshot({ path: path.join(directory, 'import.png') });
      const unchanged = await exported(page, fixture.kind, ext);
      await fs.writeFile(path.join(directory, `unchanged.${ext}`), unchanged);
      result.unchangedByteIdentical = source.equals(unchanged);
      stages.push('unchanged-export');
      expect(unchanged).toEqual(source);
      let edited = true;
      if (fixture.kind === 'word') {
        const editor = page.getByRole('textbox', { name: 'Document text', exact: true });
        await editor.click();
        await expect(editor).toBeFocused();
        await page.keyboard.press('Control+End');
        await page.keyboard.press('Enter');
        await page.keyboard.insertText(marker);
        await expect(editor).toContainText(marker);
      } else if (fixture.kind === 'powerpoint') {
        await page.getByRole('button', { name: 'Text box', exact: true }).click();
        await page.getByRole('textbox', { name: 'Object text', exact: true }).fill(marker);
      } else {
        const bar = page.getByRole('textbox', { name: 'Formula bar', exact: true });
        if (await bar.isEditable()) {
          await bar.fill(marker);
          await bar.press('Enter');
        } else {
          edited = false;
          result.editUnavailable =
            'Initially selected cell is protected; this scenario does not bypass protection.';
        }
      }
      if (edited) {
        const output = await exported(page, fixture.kind, ext);
        await fs.writeFile(path.join(directory, `edited.${ext}`), output);
        const zip = await JSZip.loadAsync(output);
        const xml = await Promise.all(
          Object.values(zip.files)
            .filter((f) => !f.dir && f.name.endsWith('.xml'))
            .map((f) => f.async('string')),
        );
        expect(xml.some((text) => text.includes(marker))).toBe(true);
        stages.push('edited-export');
        await page.getByRole('button', { name: 'Open', exact: true }).click();
        await page.locator('input[type=file][multiple]').setInputFiles({
          name: `Reimported.${ext}`,
          mimeType: 'application/octet-stream',
          buffer: output,
        });
        await expect(page.getByRole('textbox', { name: 'File name', exact: true })).toHaveValue(
          'Reimported',
          { timeout: 20000 },
        );
        stages.push('edited-reimport');
        await page.screenshot({ path: path.join(directory, 'reimport.png') });
      }
      expect(errors).toEqual([]);
      expect(external).toEqual([]);
      result.status = 'workflow-passed';
    } catch (error) {
      result.status = 'failed';
      result.error = error instanceof Error ? error.message : String(error);
      throw error;
    } finally {
      await fs.writeFile(
        path.join(directory, 'workflow.json'),
        JSON.stringify(result, null, 2) + '\n',
      );
    }
  });
}
