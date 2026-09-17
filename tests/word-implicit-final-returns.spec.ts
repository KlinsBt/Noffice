import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import reference from './fixtures/native-word-implicit-final-returns.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (data: Buffer) => createHash('sha256').update(data).digest('hex');
const run = process.env.NOFFICE_IMPLICIT_RETURN_RUN || 'returns-browser-v1';
if (!/^returns-browser-v\d+$/.test(run)) throw Error('Invalid native return run');
for (const sample of reference.rows) test(`Word native final-section return ${sample.name} remains editable after Word save`, async ({ page }) => {
  const root = `.local/word-implicit-final-sections/${run}/${sample.name}`;
  await fs.mkdir(root, { recursive: true });
  const source = await fs.readFile(sample.fixture); expect(hash(source)).toBe(sample.sourceHash);
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  const name = 'Section return ' + sample.name;
  await page.locator('input[type=file][multiple]').setInputFiles({ name: name + '.docx', buffer: source,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const body = page.getByRole('textbox', { name: 'Document text', exact: true });
  const papers = page.locator('.paper-wrap.source-layout > .paper, .section-page');
  const text = sample.snapshot.text.replace(/[\r\f]/g, '');
  await expect(body).toHaveText(text); await expect(papers).toHaveCount(sample.snapshot.pages);
  const outputs: Record<string, string> = {};
  const capture = async (stage: string) => {
    for (const [extension, label] of [['docx', 'DOCX file Editable in Microsoft Word'], ['pdf', 'PDF file']]) {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const pending = page.waitForEvent('download'); await page.getByRole('button', { name: label, exact: true }).click();
      await (await pending).saveAs(`${root}/${stage}.${extension}`);
      outputs[`${stage}.${extension}`] = hash(await fs.readFile(`${root}/${stage}.${extension}`));
    }
  };
  await capture('original'); expect(outputs['original.docx']).toBe(sample.sourceHash);
  await body.focus(); await page.keyboard.press('Control+Home'); await page.keyboard.press('End');
  await page.keyboard.type(' Browser');
  const edited = text.replace('first paragraph.', 'first paragraph. Browser');
  await expect(body).toHaveText(edited);
  await page.getByRole('button', { name: 'Undo', exact: true }).click(); await expect(body).toHaveText(text);
  await page.getByRole('button', { name: 'Redo', exact: true }).click(); await expect(body).toHaveText(edited);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await capture('edited'); await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name, exact: true }).click();
  await expect(body).toHaveText(edited); await expect(papers).toHaveCount(sample.snapshot.pages);
  await capture('reloaded');
  const originalZip = await JSZip.loadAsync(source);
  for (const stage of ['edited', 'reloaded']) {
    const zip = await JSZip.loadAsync(await fs.readFile(`${root}/${stage}.docx`));
    // Native-saved files already have a save-session register. Existing typing
    // provenance adds the new run's ID there; every unrelated setting survives.
    const settings = await page.evaluate(({ before, after, documentXml }) => {
      const parser = new DOMParser(), namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const a = parser.parseFromString(before, 'application/xml'), b = parser.parseFromString(after, 'application/xml');
      const body = parser.parseFromString(documentXml, 'application/xml');
      const entries = (doc: XMLDocument) => [...doc.getElementsByTagNameNS(namespace, 'rsid')];
      const old = new Set(entries(a).map(element => element.getAttributeNS(namespace, 'val')));
      const added = entries(b).filter(element => !old.has(element.getAttributeNS(namespace, 'val')));
      const ids = added.map(element => element.getAttributeNS(namespace, 'val'));
      const used = new Set([...body.getElementsByTagNameNS(namespace, '*')].flatMap(element =>
        [...element.attributes].filter(attribute => attribute.namespaceURI === namespace && attribute.localName.startsWith('rsid')).map(attribute => attribute.value)));
      added.forEach(element => element.remove());
      const canonical = (element: Element): unknown => [element.namespaceURI, element.tagName,
        [...element.attributes].map(attribute => [attribute.namespaceURI, attribute.name, attribute.value]).sort((a, b) => String(a).localeCompare(String(b))),
        [...element.childNodes].filter(node => node.nodeType === 1 || node.textContent?.trim()).map(node => node.nodeType === 1 ? canonical(node as Element) : node.textContent)];
      return { before: canonical(a.documentElement), after: canonical(b.documentElement), ids, registeredRuns: ids.every(id => used.has(id!)) };
    }, { before: await originalZip.file('word/settings.xml')!.async('string'),
      after: await zip.file('word/settings.xml')!.async('string'), documentXml: await zip.file('word/document.xml')!.async('string') });
    expect(settings.after).toEqual(settings.before); expect(settings.ids).toHaveLength(1);
    expect(settings.registeredRuns).toBe(true);
    for (const part of Object.keys(originalZip.files).filter(part => !originalZip.files[part].dir && !['word/document.xml', 'word/settings.xml'].includes(part)))
      expect(await zip.file(part)?.async('nodebuffer'), part).toEqual(await originalZip.file(part)!.async('nodebuffer'));
  }
  await fs.writeFile(root + '/report.json', JSON.stringify({ name: sample.name, sourceHash: sample.sourceHash, outputs,
    buildHash: await wordStoryBuildHash(), testHash: hash(await fs.readFile('tests/word-implicit-final-returns.spec.ts')) }, null, 2));
});
