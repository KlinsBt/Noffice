import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';

test('right tab fields remeasure after font edits, history, zoom and a deferred local font', async ({ page }) => {
  const zip = await JSZip.loadAsync(await fs.readFile('tests/fixtures/word-tab-stops-body.docx'));
  const xml = await zip.file('word/document.xml')!.async('string');
  const section = xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)![0];
  const paragraph = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="800" w:lineRule="exact"/>'
    + '<w:tabs><w:tab w:val="right" w:pos="4800"/></w:tabs></w:pPr>'
    + '<w:r><w:rPr><w:rFonts w:ascii="Deferred Tab Font" w:hAnsi="Deferred Tab Font"/><w:sz w:val="20"/></w:rPr>'
    + '<w:t>L</w:t><w:tab/><w:t>Wide field</w:t></w:r></w:p>';
  zip.file('word/document.xml', xml.replace(/<w:body>[\s\S]*?<\/w:body>/, '<w:body>' + paragraph + section + '</w:body>'));
  await page.setViewportSize({ width: 1500, height: 1200 });await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name: 'tab-font-cache.docx',
    buffer: await zip.generateAsync({ type: 'nodebuffer' }), mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  const geometry = () => body.locator('p').first().evaluate((p) => {
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    const rects: DOMRect[] = [];
    while (walker.nextNode()) {
      const text = walker.currentNode as Text;
      if (text.parentElement?.closest('[data-word-tab],.ProseMirror-widget')) continue;
      if (!text.data.includes('Wide field')) continue;
      const r = document.createRange();r.selectNodeContents(text);rects.push(r.getBoundingClientRect());
    }
    const scale = parseFloat(getComputedStyle(p.closest('.paper-wrap')!).zoom) || 1;
    return { right: (rects.at(-1)!.right - p.getBoundingClientRect().left) * .75 / scale,
      width: rects.reduce((sum, r) => sum + r.width, 0) * .75 / scale };
  });
  const aligned = async () => {
    await expect(body.locator('[data-word-tab-measured=true]')).toHaveCount(1);
    await expect.poll(async () => Math.abs((await geometry()).right - 240)).toBeLessThanOrEqual(.15);
  };
  await aligned();const before = await geometry();
  // A real deferred FontFace changes glyph advances while the immutable
  // paragraph and its CSS font-family string remain identical.
  const loaded = await page.evaluate(async () => {
    let events = 0;const done = () => { events++; };
    document.fonts.addEventListener('loadingdone', done);
    const face = new FontFace('Deferred Tab Font', 'local("Courier New")');document.fonts.add(face);
    await document.fonts.load('10pt "Deferred Tab Font"');await document.fonts.ready;
    await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
    document.fonts.removeEventListener('loadingdone', done);return { status: face.status, events };
  });
  expect(loaded.status).toBe('loaded');expect(loaded.events).toBeGreaterThan(0);
  await aligned();expect(Math.abs((await geometry()).width - before.width)).toBeGreaterThan(1);
  await body.focus();await page.keyboard.press('Control+a');
  const size = page.getByRole('spinbutton', { name: 'Font size', exact: true });
  await size.fill('20');await size.press('Tab');await aligned();const large = await geometry();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();await aligned();
  expect(large.width).toBeGreaterThan((await geometry()).width * 1.9);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();await aligned();
  for (const action of ['Zoom out', 'Zoom out', 'Zoom in', 'Zoom in']) {
    await page.getByRole('button', { name: action, exact: true }).click();await aligned();
  }
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'tab-font-cache', exact: true }).click();await aligned();
});
