import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';

test('tab measurement stops before over-budget field work and preserves editing recovery', async ({ page }) => {
  test.setTimeout(90000);
  const root = '.local/word-tab-budget';await fs.mkdir(root, { recursive: true });
  const zip = await JSZip.loadAsync(await fs.readFile('tests/fixtures/word-tab-stops-body.docx'));
  const xml = await zip.file('word/document.xml')!.async('string');
  const section = xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)![0];
  const paragraph = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="400" w:lineRule="exact"/></w:pPr>'
    + '<w:r><w:tab/><w:t>Budget</w:t></w:r>'.repeat(1001) + '</w:p>';
  zip.file('word/document.xml', xml.replace(/<w:body>[\s\S]*?<\/w:body>/, '<w:body>' + paragraph + section + '</w:body>'));
  const input = await zip.generateAsync({ type: 'nodebuffer' });await fs.writeFile(root + '/input.docx', input);
  const errors: string[] = [];page.on('pageerror', (e) => errors.push(e.message));
  const timings: Record<string, number> = {};
  await page.addInitScript(() => {
    const evidence = { frames: [] as number[], pending: 0 };
    (window as unknown as { tabBudgetEvidence: typeof evidence }).tabBudgetEvidence = evidence;
    const rect = HTMLElement.prototype.getBoundingClientRect;
    HTMLElement.prototype.getBoundingClientRect = function () {
      // Observe a synchronous measurement batch, including work invoked outside
      // requestAnimationFrame. The real rectangles and keyboard timing remain.
      if (this.tagName === 'SPAN' && this.parentElement?.matches('body,[data-word-tab-probes]') && this.style.position === 'fixed'
        && this.style.left === '-100000px' && this.style.width === 'max-content' && this.textContent === 'Budget') {
        if (evidence.pending++ === 0) queueMicrotask(() => { evidence.frames.push(evidence.pending); evidence.pending = 0; });
      }
      return rect.call(this);
    };
  });
  await page.setViewportSize({ width: 1500, height: 1200 });await page.goto('/');
  const profiler = process.env.NOFFICE_TAB_PROFILE === '1' ? await page.context().newCDPSession(page) : null;
  if (profiler) { await profiler.send('Profiler.enable');await profiler.send('Profiler.start'); }
  const opening = performance.now();
  await page.locator('input[type=file][multiple]').setInputFiles({ name: 'tab-budget.docx', buffer: input, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  await expect(body.locator('[data-word-tab-measured=true]')).toHaveCount(1000);
  await expect(body.locator('[data-word-tab-measured=false]')).toHaveCount(1);
  const counts = () => page.evaluate(() => (window as unknown as { tabBudgetEvidence: { frames: number[] } }).tabBudgetEvidence.frames);
  await fs.writeFile(root + '/measurement-diagnostic.json', JSON.stringify(await page.evaluate(() => (window as unknown as { tabBudgetEvidence: unknown }).tabBudgetEvidence), null, 2));
  await expect.poll(async () => (await counts()).length).toBeGreaterThan(0);
  const before = await counts();expect(Math.max(...before)).toBeLessThanOrEqual(1000);
  // The initially measured prefix is not proof that this wrapped paragraph has
  // converged. Exhausted layout must reject every uncertain spacer, settle and
  // allow recovery, instead of continuously repainting approximate positions.
  await expect(body.locator('[data-word-tab-measured=false]')).toHaveCount(1001);
  await expect(body.locator('[data-word-baseline-paint]')).toHaveCount(0);
  const stability = async () => body.evaluate(async (host) => {
    const states: string[] = [];
    const sample: { first: string[]; last: string[] }[] = [];
    for (let i = 0; i < 10; i++) {
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
      states.push(host.innerHTML);
      const styles = [...host.querySelectorAll('[data-word-tab]')].map((e) => e.getAttribute('style') || '');
      sample.push({ first: styles.slice(0, 8), last: styles.slice(-8) });
    }
    return { unique: new Set(states).size, sample };
  });
  let stable = 0;
  await expect.poll(async () => {
    const result = await stability();stable = result.unique;
    await fs.writeFile(root + '/stability-diagnostic.json', JSON.stringify(result, null, 2));return stable;
  }).toBe(1);
  timings.openUntilMeasuredMs = performance.now() - opening;
  if (profiler) { const result = await profiler.send('Profiler.stop');await fs.writeFile(root + '/cpu-profile.json', JSON.stringify(result));await profiler.detach(); }
  let unexpectedDownloads = 0;const downloaded = () => { unexpectedDownloads++; };page.on('download', downloaded);
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  await page.getByRole('button', { name: 'PDF file', exact: true }).click();
  const failure = page.getByRole('alert').filter({ hasText: 'Export failed:' });
  await expect(failure).toBeVisible();const pdfFailure = await failure.textContent();
  expect(pdfFailure).toContain('PDF export is not yet available for this layout.');
  expect(unexpectedDownloads).toBe(0);page.off('download', downloaded);
  await page.getByRole('button', { name: 'Dismiss error', exact: true }).click();
  const download = async (file: string, backup = false) => {
    const started = performance.now();
    await page.getByRole('button', { name: 'Export', exact: true }).click();const waiting = page.waitForEvent('download');
    await page.getByRole('button', { name: backup ? /^Noffice backup/ : 'DOCX file Editable in Microsoft Word', exact: !backup }).click();
    await expect(page.getByRole('alert').filter({ hasText: 'Export failed:' })).toHaveCount(0);
    await (await waiting).saveAs(root + '/' + file);timings[file] = performance.now() - started;return fs.readFile(root + '/' + file);
  };
  expect(await download('original.docx')).toEqual(input);
  let started = performance.now();await body.focus();await page.keyboard.press('Control+a');await page.keyboard.insertText('Recovered');
  await expect(body).toHaveText('Recovered');timings.replaceMs = performance.now() - started;
  started = performance.now();await page.keyboard.press('Control+z');await expect(body.locator('[data-word-tab]')).toHaveCount(1001);timings.undoMs = performance.now() - started;
  started = performance.now();await page.keyboard.press('Control+y');await expect(body).toHaveText('Recovered');timings.redoMs = performance.now() - started;
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const edited = await download('edited.docx');const editedZip = await JSZip.loadAsync(edited);
  expect(await editedZip.file('word/document.xml')!.async('string')).toContain('Recovered');
  const backup = JSON.parse((await download('recovery.noffice', true)).toString('utf8'));
  expect(Buffer.from(backup.original.base64, 'base64')).toEqual(input);
  const after = await counts();expect(Math.max(...after)).toBeLessThanOrEqual(1000);
  await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'tab-budget', exact: true }).click();await expect(body).toHaveText('Recovered');
  expect(errors).toEqual([]);
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  await fs.writeFile(root + '/browser-report.json', JSON.stringify({ sourceHash: hash(input), editedHash: hash(edited), before, after, stable, pdfFailure, errors, timings,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-tab-budget.spec.ts')) }, null, 2));
});

test('native tab-budget recovery returns through further editing and reload', async ({ page }) => {
  test.skip(process.env.NOFFICE_TAB_BUDGET_NATIVE_RETURN !== '1', 'Requires independently edited installed-Word recovery output');
  const root = '.local/word-tab-budget', run = process.env.NOFFICE_TAB_BUDGET_NATIVE_RUN || 'native-v1';
  const bytes = await fs.readFile(`${root}/${run}/returned.docx`);
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles({ name: 'native-tab-budget.docx', buffer: bytes, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });await expect(body).toHaveText('Recovered Native');
  await body.focus();await page.keyboard.press('Control+End');await page.keyboard.insertText(' Browser');await expect(body).toHaveText('Recovered Native Browser');
  await page.keyboard.press('Control+z');await expect(body).toHaveText('Recovered Native');await page.keyboard.press('Control+y');await expect(body).toHaveText('Recovered Native Browser');
  await page.keyboard.press('Control+z');await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();await page.getByRole('button', { name: 'native-tab-budget', exact: true }).click();
  await expect(body).toHaveText('Recovered Native');
  await page.getByRole('button', { name: 'Export', exact: true }).click();const waiting = page.waitForEvent('download');
  await page.getByRole('button', { name: 'DOCX file Editable in Microsoft Word', exact: true }).click();await (await waiting).saveAs(root + '/native-return.docx');
  expect(await fs.readFile(root + '/native-return.docx')).toEqual(bytes);
  const hash = (b: Buffer) => createHash('sha256').update(b).digest('hex');
  await fs.writeFile(root + '/native-return-report.json', JSON.stringify({ sourceHash: hash(bytes), nativeReceiptHash: hash(await fs.readFile(`${root}/${run}/native-report.json`)), buildHash: await wordStoryBuildHash() }, null, 2));
});
