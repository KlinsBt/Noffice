import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-story-paragraph-layout.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) for (const state of row.states)
  test(`Word story paragraph import ${row.name} ${state.stage}`, async ({ page }) => {
    const run = process.env.NOFFICE_STORY_PARAGRAPH_RUN || 'import-v1';
    if (!/^import-v\d+$/.test(run)) throw Error('Invalid paragraph-layout run');
    const root = `.local/word-story-paragraph-layout/${run}/${row.name}/${state.stage}`;
    await fs.mkdir(root, { recursive: true });
    const hash = (value: Buffer) => createHash('sha256').update(value).digest('hex');
    expect(hash(await fs.readFile(state.fixture))).toBe(state.docxHash);
    const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
    await page.context().grantPermissions(['local-fonts']); await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles(state.fixture);
    await expect(page.locator('.section-page')).toHaveCount(state.pages);
    // The reading/print line copies must retain the actual editor's horizontal
    // positions. This catches lost alignment/hanging spaces separately from
    // the stricter independent native glyph and line-identity comparison.
    const paint = await page.evaluate(() => {
      const original = [...document.querySelectorAll<HTMLElement>('.word-story-measure-host p')]
        .find(p => p.textContent?.startsWith('Alpha'))!;
      const id = original.getAttribute('data-source-paragraph');
      const copies = [...document.querySelectorAll<HTMLElement>('.section-page')][0]
        .querySelectorAll<HTMLElement>(`p[data-source-paragraph="${id}"]`);
      const read = (p: HTMLElement) => {
        const bounds = p.getBoundingClientRect(), scale = bounds.width / parseFloat(getComputedStyle(p).width);
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT), glyphs: { text: string; x: number }[] = [];
        let n: Node | null;
        while ((n = walker.nextNode())) {
          const excluded = n.parentElement?.closest('.ProseMirror-widget,[contenteditable=false]');
          if (excluded && p.contains(excluded)) continue;
          for (let i = 0; i < (n.textContent?.length || 0); i++) {
            const r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + 1);
            glyphs.push({ text: n.textContent![i], x: (r.getBoundingClientRect().left - bounds.left) / scale });
          }
        }
        return glyphs;
      };
      const expected = read(original), actual = [...copies].flatMap(read);
      return { expectedText: expected.map(g => g.text).join(''), actualText: actual.map(g => g.text).join(''),
        delta: expected.length === actual.length ? Math.max(...expected.map((g, i) => /\s/.test(g.text) ? 0 : Math.abs(g.x - actual[i].x))) : null };
    });
    expect(paint.actualText).toBe(paint.expectedText); expect(paint.delta).not.toBeNull(); expect(paint.delta!).toBeLessThanOrEqual(.04);
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: 'PDF file', exact: true }).click();
    await (await waiting).saveAs(root + '/actual.pdf');
    expect(errors).toEqual([]);
    await fs.writeFile(root + '/browser-report.json', JSON.stringify({ name: row.name, stage: state.stage,
      sourceHash: state.docxHash, pdfHash: hash(await fs.readFile(root + '/actual.pdf')), errors, paint,
      buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-story-paragraph-layout-import.spec.ts')) }, null, 2));
  });
