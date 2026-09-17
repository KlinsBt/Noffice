import { it, expect, vi } from 'vitest';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { readDocx, wordXml, descendants } from './docx-import';
import { docxStyles } from './docx-styles';
import { docxTabStops, docxDefaultTab, docxTabSettings } from './docx-tabs';
import { wordJSON } from './word-extensions';
import { wordTabStops, wordDefaultTab, wordDecimalSymbol } from './word-tab-stops';
import reference from '../tests/fixtures/native-word-tab-stops.json';
import { newFile } from './model';
import { contentFingerprint } from './office-preservation';
import { hydrateWordStructure, needsWordStructure } from './word-structure';
import { changedWordStoryCreation } from './word-story-create';
import { authoredWordStoryTemplates } from './word-authored-stories';
import tabGlyphs from '../tests/fixtures/native-word-tab-stop-glyphs.json';
import { wordNativeBaseline } from './word-native-baseline';
import JSZip from 'jszip';
import { exportRetainedDocument } from './docx-preserve';
import lineBoxes from '../tests/fixtures/native-word-tab-line-boxes.json';

it('binds the exact20 Arial10 baseline to all 21 native tab paragraphs', () => {
  const baseline = wordNativeBaseline(10 / .75, 11.5 / .75, 1.1499, 'Arial', '400', 'normal', 'normal', { rule: 'exact', line: 400 });
  expect(baseline).not.toBeNull();
  const offsets = lineBoxes.rows.flatMap((row) => row.offsets.map((item) => item.offset));
  expect(offsets).toHaveLength(91);
  for (const row of lineBoxes.rows) expect(row.glyphMatricesUnchanged).toBe(true);
  const midpoint = (Math.min(...offsets) + Math.max(...offsets)) / 2;
  expect(baseline! * .75).toBe(Math.round(midpoint * 20) / 20);
  for (const offset of offsets) expect(Math.abs(offset - baseline! * .75)).toBeLessThanOrEqual(.15);
  const rows = tabGlyphs.rows.filter((row) => row.document === 'body');
  expect(rows).toHaveLength(21);
  for (const row of rows)
    expect(Math.abs(row.glyphs[0].y - (72 + row.paragraph * 20 + baseline! * .75)), row.name).toBeLessThanOrEqual(.15);
});

vi.mock('mammoth', async (original) => {
  const actual = await original<typeof import('mammoth')>();
  return {
    ...actual,
    convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
      actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]),
  };
});
const alignments = ['left', 'center', 'right', 'decimal', 'bar'];
const leaders = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'];
const nativeStops = (tabs: { position: number; alignment: number; leader: number; custom: boolean }[]) =>
  tabs.filter((t) => t.custom).map((t) => ({ position: Math.round(t.position * 20), alignment: alignments[t.alignment], leader: leaders[t.leader] }));

for (const sample of reference.rows)
  it(`retains the native ${sample.name} tab cascade through the editable schema`, async () => {
    const bytes = await readFile('tests/fixtures/' + sample.file);
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(sample.fixtureHash);
    const { content } = await readDocx(Uint8Array.from(bytes).buffer);
    const paragraphs = wordJSON(content.html).content!.filter((p) => p.type === 'paragraph');
    for (const [index, p] of sample.state.paragraphs.entries()) {
      const attrs = paragraphs[index].attrs!;
      expect(attrs.paragraphTabs || [], `${sample.name} paragraph ${index}`).toEqual(nativeStops(p.tabs));
      expect(attrs.paragraphDefaultTab).toBe(Math.round(sample.state.defaultTab * 20));
      expect(attrs.paragraphDecimalSymbol).toBe(',');
    }
    if (sample.name === 'stories')
      for (const story of sample.state.stories) {
        const part = content.stories!.parts.find((p) => p.kind === (story.kind === 'Headers' ? 'header' : 'footer'))!;
        expect(wordJSON(part.html).content![0].attrs!.paragraphTabs).toEqual(nativeStops(story.tabs));
      }
    for (const template of Object.values(content.stories!.emptyTemplates!))
      expect(wordJSON(template!).content![0].attrs!.paragraphTabs).toEqual([
        { position: 4536, alignment: 'center', leader: 'none' },
        { position: 9072, alignment: 'right', leader: 'none' },
      ]);
  });

it('merges tabs by position and preserves inheritance through an empty tab list', () => {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const styles = `<w:styles xmlns:w="${w}"><w:style w:styleId="base" w:type="paragraph"><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:pos="1440"/></w:tabs></w:pPr></w:style><w:style w:styleId="child" w:type="paragraph"><w:basedOn w:val="base"/><w:pPr><w:tabs/></w:pPr></w:style></w:styles>`;
  const source = wordXml(`<w:document xmlns:w="${w}"><w:body><w:p><w:pPr><w:pStyle w:val="child"/><w:tabs><w:tab w:val="clear" w:pos="+0720"/><w:tab w:val="center" w:pos="2160"/></w:tabs></w:pPr></w:p></w:body></w:document>`);
  const p = descendants(source, 'p')[0];
  expect(docxTabStops(docxStyles(styles)(p))).toEqual([
    { position: 1440, alignment: 'right', leader: 'none' },
    { position: 2160, alignment: 'center', leader: 'none' },
  ]);
});

const blobBytes = (blob: Blob) => new Promise<ArrayBuffer>((resolve, reject) => {
  const reader = new FileReader(); reader.onload = () => resolve(reader.result as ArrayBuffer);
  reader.onerror = () => reject(reader.error); reader.readAsArrayBuffer(blob);
});

it('exports changed inherited stops without reviving earlier clears or changing unrelated parts', async () => {
  const data = Uint8Array.from(await readFile('tests/fixtures/word-tab-stops-body.docx')).buffer;
  const content = (await readDocx(data)).content;
  const file = newFile('word', 'Tab edits', content); file.original = { name: 'source.docx', data };
  const html = new DOMParser().parseFromString(content.html, 'text/html');
  const changes = [
    { index: 17, stops: [{ position: 2880, alignment: 'center', leader: 'dot' }] },
    { index: 18, stops: [{ position: 4320, alignment: 'right', leader: 'none' }] },
    { index: 20, stops: [{ position: 2880, alignment: 'decimal', leader: 'none' }] },
  ];
  for (const { index, stops } of changes) html.querySelectorAll('p')[index].dataset.wordTabs = JSON.stringify(stops);
  if (file.content.kind !== 'word') throw Error('Expected Word');
  file.content.html = html.body.innerHTML;
  const output = await blobBytes(await exportRetainedDocument(file));
  const reopened = await readDocx(output), nodes = wordJSON(reopened.content.html).content!;
  for (const { index, stops } of changes) expect(nodes[index].attrs!.paragraphTabs).toEqual(stops);
  const before = await JSZip.loadAsync(data), after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'word/document.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(await before.file(path)!.async('uint8array'));
  expect(file.original!.data).toBe(data);
});

it('rejects paragraph metadata that would silently change document-wide tab settings', async () => {
  const data = Uint8Array.from(await readFile('tests/fixtures/word-tab-stops-body.docx')).buffer;
  for (const [attribute, value, error] of [
    ['data-word-default-tab', '360', 'default tab interval'],
    ['data-word-decimal-symbol', '.', 'decimal-tab separator'],
  ]) {
    const content = (await readDocx(data)).content;
    const file = newFile('word', 'Invalid tab settings', content); file.original = { name: 'source.docx', data };
    const html = new DOMParser().parseFromString(content.html, 'text/html');
    html.querySelector('p')!.setAttribute(attribute, value);
    if (file.content.kind !== 'word') throw Error('Expected Word');
    file.content.html = html.body.innerHTML;
    await expect(exportRetainedDocument(file)).rejects.toThrow(error);
    expect(file.original!.data).toBe(data);
  }
});

it('bounds untrusted tab definitions without silently accepting malformed positions', () => {
  const stop = { position: 720, alignment: 'left', leader: 'none' };
  for (const value of ['{', [stop, stop], [{ ...stop, position: 1.2 }], [{ ...stop, alignment: 'script' }], Array(257).fill(stop)])
    expect(wordTabStops(value)).toBeNull();
  expect(wordDefaultTab('-1')).toBeNull();
  expect(wordDefaultTab('32768')).toBeNull();
  expect(wordDefaultTab('0')).toBe(0);
  expect(docxDefaultTab()).toBe(720);
  expect(wordDecimalSymbol(',')).toBe(',');
  expect(wordDecimalSymbol('\n')).toBeNull();
});

it.each([
  '<w:tab w:val="clear"/>',
  '<w:tab w:val="clear" w:pos="invalid"/>',
  '<w:tab w:val="clear" w:pos="999999"/>',
  '<w:tab w:val="left" w:pos="720" w:leader="unknown"/>',
  '<w:tab w:val="left" w:pos="720" w:unknown="1"/>',
  '<w:unknown/>',
  '<w:tab w:val="left" w:pos="720"/><w:tab w:val="clear" w:pos="720"/>',
])('keeps an invalid inherited tab definition visible after clearing or overriding: %s', (invalid) => {
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const styles = `<w:styles xmlns:w="${ns}"><w:style w:type="paragraph" w:styleId="bad"><w:pPr><w:tabs>${invalid}</w:tabs></w:pPr></w:style></w:styles>`;
  const doc = wordXml(`<w:document xmlns:w="${ns}"><w:body><w:p><w:pPr><w:pStyle w:val="bad"/><w:tabs><w:tab w:val="clear" w:pos="720"/><w:tab w:val="right" w:pos="1440"/></w:tabs></w:pPr></w:p></w:body></w:document>`);
  expect(docxTabStops(docxStyles(styles)(descendants(doc, 'p')[0]))).toBeNull();
});

it('distinguishes omitted tab settings from present empty or malformed settings', () => {
  const settings = (body: string) => wordXml(`<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${body}</w:settings>`).documentElement;
  expect(docxTabSettings(settings(''))).toEqual({ interval: 720, decimal: '', decimalInvalid: false });
  expect(docxTabSettings(settings('<w:defaultTabStop/>')).interval).toBeNull();
  expect(docxTabSettings(settings('<w:defaultTabStop w:val="+00000720"/>')).interval).toBe(720);
  for (const raw of ['', 'ab', ' ']) expect(docxTabSettings(settings(`<w:decimalSymbol w:val="${raw}"/>`)).decimalInvalid).toBe(true);
  expect(docxTabSettings(settings('<w:decimalSymbol w:val=","/>')).decimalInvalid).toBe(false);
});

it.each(['defaultTabStop', 'decimalSymbol'])('preserves original bytes and warns on a present empty %s', async (setting) => {
  const zip = await JSZip.loadAsync(await readFile('tests/fixtures/word-tab-stops-stories.docx'));
  zip.file('word/settings.xml', (await zip.file('word/settings.xml')!.async('string')).replace(new RegExp(`<w:${setting}\\b[^>]*/>`), `<w:${setting}/>`));
  const data = await zip.generateAsync({ type: 'arraybuffer' });const read = await readDocx(data);
  expect(read.messages.some((m) => m.message.includes('tab definitions'))).toBe(true);
  expect(wordJSON(read.content.html).content![0].attrs!.paragraphTabUnsupported).toBe(true);
  if (setting === 'decimalSymbol') expect(wordJSON(read.content.html).content![0].attrs!.paragraphDefaultTab).toBe(720);
  expect(wordJSON(read.content.stories!.parts[0].html).content![0].attrs!.paragraphTabUnsupported).toBe(true);
  for (const template of Object.values(read.content.stories!.emptyTemplates!))
    expect(wordJSON(template!).content![0].attrs!.paragraphTabUnsupported).toBe(true);
  const file = newFile('word', 'Preserved invalid tabs', read.content);file.original = { name: 'original.docx', data };
  expect(await blobBytes(await exportRetainedDocument(file))).toEqual(data);
});

function omitTabs(html: string) {
  return html.replace(/ data-word-(?:tabs|default-tab|decimal-symbol|tab-layout-unsupported)="[^"]*"/g, '');
}

for (const edited of [false, true])
  it(`recovers omitted body/story tab metadata while preserving ${edited ? 'edited' : 'unchanged'} legacy content`, async () => {
    const bytes = Uint8Array.from(await readFile('tests/fixtures/word-tab-stops-stories.docx')).buffer;
    const source = (await readDocx(bytes)).content;
    const content = {
      ...source, tabStopsVersion: undefined,
      html: omitTabs(source.html),
      stories: { ...source.stories!, parts: source.stories!.parts.map((p) => ({ ...p, html: omitTabs(p.html) })) },
    };
    const fingerprint = await contentFingerprint(content);
    if (edited) {
      const story = new DOMParser().parseFromString(content.stories.parts[0].html, 'text/html');
      story.querySelector('p')!.append(' Edited');
      content.stories.parts[0].html = story.body.innerHTML;
      expect(await contentFingerprint(content)).not.toBe(fingerprint);
    }
    const file = { ...newFile('word', 'Legacy tabs', content), revision: 7,
      original: { name: 'source.docx', data: bytes, contentFingerprint: fingerprint } };
    expect(needsWordStructure(file)).toBe(true);
    const next = await hydrateWordStructure(file);
    expect(next.revision).toBe(7);
    expect(next.original!.data).toBe(bytes);
    expect(next.content.kind).toBe('word');
    if (next.content.kind !== 'word') throw Error('Expected Word');
    expect(wordJSON(next.content.html)).toEqual(wordJSON(source.html));
    expect(omitTabs(next.content.stories!.parts[0].html)).toBe(content.stories.parts[0].html);
    expect(wordJSON(next.content.stories!.parts[0].html).content![0].attrs!.paragraphTabs).toHaveLength(2);
    expect(await hydrateWordStructure(next)).toBe(next);
    expect(needsWordStructure(next)).toBe(false);
    expect(next.original!.contentFingerprint).toBe(edited ? fingerprint : await contentFingerprint(next.content));
  });

it('recovers fresh authored story tabs without inventing retained body identities', async () => {
  const file = newFile('word');
  if (file.content.kind !== 'word') throw Error('Expected Word');
  const content = file.content;
  content.stories = changedWordStoryCreation(content, 'authored-body', 'header', 'default',
    authoredWordStoryTemplates(content)!.header.replace('</p>', 'Fresh header</p>'));
  content.tabStopsVersion = undefined;
  content.stories.parts = content.stories.parts.map((p) => ({ ...p, html: omitTabs(p.html) }));
  const next = await hydrateWordStructure(file);
  if (next.content.kind !== 'word') throw Error('Expected Word');
  expect(next.original).toBeUndefined();
  expect(next.content.docxStructure).toBeUndefined();
  expect(next.content.html).not.toContain('data-source-paragraph');
  expect(wordJSON(next.content.stories!.parts[0].html).content![0].attrs!.paragraphTabs).toHaveLength(2);
  expect(needsWordStructure(next)).toBe(false);
});
