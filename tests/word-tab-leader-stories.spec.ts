import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-tab-leader-stories.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const native of reference.rows) test(`${native.name} header/footer leaders retain native cells through editing and exports`, async ({ page }) => {
  test.setTimeout(90000);
  const run = process.env.NOFFICE_LEADER_STORIES_RUN || 'browser-v1';
  expect(run).toMatch(/^browser-v\d+$/);
  const root = `.local/word-tab-leader-stories/sources-v2/${native.name}/${run}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile(`tests/fixtures/word-tab-leader-story-${native.name}.docx`);
  expect(hash(source)).toBe(native.sourceHash);
  const outputs: Record<string, string> = {}, errors: string[] = [], checks: unknown[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);await page.setViewportSize({ width: 1500, height: 1200 });
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles({ name: `leader-story-${native.name}.docx`,
    buffer: source, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const check = async (stage: 'source' | 'expected') => {
    await expect(page.locator('.section-page')).toHaveCount(2);
    await expect(page.locator('.word-page-story')).toHaveCount(4);
    await expect(page.locator('.word-page-story[data-word-story*="header"] p').first())
      .toHaveText(stage === 'source' ? 'A\tB' : 'Edited\tB');
    const actual = await page.locator('.section-page').evaluateAll((pages) => pages.flatMap((page, index) => {
      const scale = parseFloat(getComputedStyle(page.closest('.paper-wrap')!).zoom) || 1;
      return [...page.querySelectorAll<HTMLElement>('.word-page-story')].flatMap((story) =>
        [...story.querySelectorAll('p')].map((p, number) => {
          const semantic = p.querySelector<HTMLElement>('[data-word-tab]');
          const tab = semantic?.closest<HTMLElement>('[data-word-tab-measured]') || semantic;
          const pseudo = tab ? getComputedStyle(tab, '::after') : null;
          return { page: index + 1, kind: story.dataset.wordStory?.includes('header') ? 'header' : 'footer', index: number,
            painted: tab?.dataset.wordTabLeaderPainted, offsets: JSON.parse(tab?.dataset.wordTabLeaderOffsets || 'null') as number[] | null,
            x: tab ? (tab.getBoundingClientRect().left - page.getBoundingClientRect().left) * .75 / scale : null,
            color: pseudo?.color, glyph: pseudo?.content, size: pseudo ? parseFloat(pseudo.fontSize) * .75 : null };
        }));
    }));
    expect(actual).toHaveLength(20);
    const expected = native.states.find((s) => s.stage === stage)!;
    for (const row of actual) {
      const sample = expected.pages[row.page - 1].cases.find((c) => c.kind === row.kind && c.index === row.index)!;
      expect(row.painted, `${row.kind}${row.index}`).toBe('true');
      expect(row.offsets).toHaveLength(sample.leaders.length);
      expect(Math.max(0, ...row.offsets!.map((offset, i) => Math.abs(row.x! + offset * .75 - sample.leaders[i].x)))).toBeLessThanOrEqual(.15);
      expect(row.glyph).toBe(JSON.stringify(sample.leaders[0]?.text || ''));
      expect(row.color).toBe(`rgb(${sample.leaders[0].color.slice(0, 3).join(', ')})`);
      expect(Math.abs(row.size! - sample.leaders[0].size)).toBeLessThanOrEqual(.15);
    }
    checks.push({ stage, actual });
  };
  const download = async (name: string, label = 'PDF file') => {
    await page.getByRole('button', { name: 'Export', exact: true }).click();const pending = page.waitForEvent('download');
    await page.getByRole('button', { name: label, exact: true }).click();await (await pending).saveAs(root + '/' + name);
    outputs[name] = hash(await fs.readFile(root + '/' + name));
  };
  await expect(page.locator('.word-page-story [data-word-tab-leader-painted=true]')).toHaveCount(20);
  await check('source');await download('source.pdf');
  await page.getByRole('button', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
  await page.getByRole('button', { name: /^Section 1 \u2014 Default header/ }).click();
  const editor = page.getByRole('textbox', { name: 'Header or footer text', exact: true });
  await expect(editor.locator('[data-word-tab-leader-painted=true]')).toHaveCount(5);
  await editor.focus();await page.keyboard.press('Control+Home');await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.insertText('Edited');await expect(editor.locator('p').first()).toHaveText('Edited\tB');
  await page.keyboard.press('Control+z');await expect(editor.locator('p').first()).toHaveText('A\tB');
  await page.keyboard.press('Control+y');await expect(editor.locator('p').first()).toHaveText('Edited\tB');
  await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
  await expect(page.locator('.word-page-story [data-word-tab-leader-painted=true]')).toHaveCount(20);
  await check('expected');
  await page.getByRole('button', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();await check('source');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();await check('expected');
  await download('edited.pdf');await download('edited.docx', 'DOCX file Editable in Microsoft Word');
  await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: `leader-story-${native.name}`, exact: true }).click();
  await expect(page.locator('.word-page-story [data-word-tab-leader-painted=true]')).toHaveCount(20);
  await check('expected');await download('reloaded.pdf');
  expect(errors).toEqual([]);
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(source), nativeReceiptHash: native.nativeReceiptHash,
    outputs, errors, checks, buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-tab-leader-stories.spec.ts')) }, null, 2));
});
