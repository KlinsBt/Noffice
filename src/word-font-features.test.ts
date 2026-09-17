import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { readDocx, wordXml, WORD_NS } from './docx-import';
import { wordExtensions, wordJSON } from './word-extensions';
import { wordFontFeaturesValue, wordFontFeaturesXml, WORD_2010_NS } from './word-font-features';
import { hydrateWordFontFeatures } from './word-font-feature-migration';
import { exportOffice, sanitizeHTML } from './formats';
import { newFile, type WordContent } from './model';
import { clonedStoryHtml } from './word-story-provenance';
import native from '../tests/fixtures/native-word-ligatures.json';

vi.mock('mammoth', async (original) => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});

it.each(native.rows)('imports native ligature setting $name through the complete style cascade', async (row) => {
  const bytes = readFileSync(`tests/fixtures/word-ligatures-${row.name}.docx`);
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(row.docxHash);
  const { content } = await readDocx(Uint8Array.from(bytes).buffer);
  const footer = content.stories!.parts.find((part) => part.html.includes('office'))!;
  const first = wordJSON(sanitizeHTML(footer.html)).content![0];
  expect(first.attrs?.paragraphFontFeatures).toBe(row.markFlags | (row.markAlternates ? 16 : 0));
  expect(first.content!.filter((node) => node.type === 'text').map((node) =>
    node.marks!.find((mark) => mark.type === 'textStyle')!.attrs!.wordFontFeatures))
    .toEqual(expect.arrayContaining([row.flags]));
  for (const node of first.content!) if (node.type === 'text')
    expect(node.marks!.find((mark) => mark.type === 'textStyle')!.attrs!.wordFontFeatures).toBe(row.flags);
  expect(content.fontFeaturesVersion).toBe(1);
});

it('validates extension namespaces, explicit zero, alternate switches and invalid settings', () => {
  const read = (body: string) => wordFontFeaturesXml(wordXml(`<w:rPr xmlns:w="${WORD_NS}" xmlns:f="${WORD_2010_NS}">${body}</w:rPr>`).documentElement);
  expect(read('')).toBe(0);
  expect(read('<w:ligatures w:val="all"/>')).toBe(0);
  expect(read('<f:ligatures f:val="none"/><f:cntxtAlts/>')).toBe(16);
  expect(read('<f:ligatures f:val="all"/><f:cntxtAlts f:val="false"/>')).toBe(15);
  for (const body of ['<f:ligatures/>', '<f:ligatures f:val="invalid"/>', '<f:cntxtAlts f:val="2"/>'])
    expect(() => read(body)).toThrow(/invalid font-feature/);
  for (const value of [null, undefined, true, -1, 32, 1.5, '01', '1px', {}, Infinity])
    expect(wordFontFeaturesValue(value)).toBeNull();
});

it('retains metadata-only spans, explicit overrides and empty-paragraph typing in ordinary history', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p data-word-paragraph-font-features="3"></p>' });
  try {
    editor.commands.insertContent('office');
    expect(editor.state.doc.firstChild!.firstChild!.marks.find((m) => m.type.name === 'textStyle')!.attrs.wordFontFeatures).toBe(3);
    editor.commands.undo(); expect(editor.getText()).toBe('');
    editor.commands.redo(); expect(editor.getText()).toBe('office');
    const html = editor.getHTML(); editor.commands.setContent(sanitizeHTML(html));
    expect(editor.getHTML()).toBe(html);
    editor.commands.setContent('<p><span data-word-font-features="3">ti<span data-word-font-features="0">fi</span></span></p>');
    expect(editor.getJSON().content![0].content!.map((node) => node.marks![0].attrs!.wordFontFeatures)).toEqual([3, 0]);
  } finally { editor.destroy(); }
});

it('writes fresh run and empty paragraph-mark feature settings into the actual DOCX', async () => {
  const file = newFile('word', 'Features');
  if (file.content.kind !== 'word') throw Error('Word fixture');
  file.content.html = '<p data-word-paragraph-font-features="3"><span data-word-font-features="1">office</span><span data-word-font-features="0"> affinity</span></p><p data-word-paragraph-font-features="15"></p>';
  const blob = await exportOffice(file);
  const zip = await JSZip.loadAsync(blob), xml = wordXml(await zip.file('word/document.xml')!.async('string'));
  expect(xml.getElementsByTagNameNS(WORD_2010_NS, 'ligatures').length).toBe(4);
  const restored = (await readDocx(await zip.generateAsync({ type: 'arraybuffer' }))).content, json = wordJSON(restored.html);
  expect(json.content!.map((p) => p.attrs!.paragraphFontFeatures)).toEqual([3, 15]);
  expect(json.content![0].content!.map((node) => node.marks!.find((m) => m.type === 'textStyle')!.attrs!.wordFontFeatures)).toEqual([1, 0]);
});

const content = (html: string): WordContent => ({ kind: 'word', html, paper: 'a4', margin: 'normal' });
it('recovers edited cloned and created story features from retained provenance', () => {
  const source = content('<p data-source-paragraph="0:0"></p>');
  const original = { path: 'word/footer.xml', kind: 'footer' as const, relationshipIds: ['rFooter'],
    html: '<p data-source-paragraph="3:0" data-word-paragraph-font-features="3"><span data-word-font-features="3">Native</span></p>' };
  const template = '<p data-source-paragraph="empty:0" data-word-paragraph-font-features="1"></p>';
  source.stories = { version: 1, evenAndOddHeaders: false, parts: [original], emptyTemplates: { header: null, footer: template } };
  const missing = (html: string) => html.replace(/ data-word-(?:paragraph-)?font-features="\d+"/g, '');
  const clonePath = 'word/footerCopy.xml', newPath = 'word/footerCreated.xml';
  const saved: WordContent = { ...source, stories: { ...source.stories, parts: [
    { ...original, path: clonePath, copiedFrom: original.path,
      html: missing(clonedStoryHtml(original, clonePath)).replace('Native', 'Edited') },
    { path: newPath, kind: 'footer', relationshipIds: ['rNew'], created: true,
      html: missing(clonedStoryHtml({ ...original, html: template }, newPath)).replace('</p>', 'New</p>') },
  ] } };
  const unchanged = structuredClone(saved), result = hydrateWordFontFeatures(saved, source);
  expect(result.stories!.parts.map((p) => wordJSON(p.html).content![0].content![0].marks!
    .find((m) => m.type === 'textStyle')!.attrs!.wordFontFeatures)).toEqual([3, 1]);
  expect(saved).toEqual(unchanged);
});
it('migrates uniform edited and unchanged mixed legacy text without guessing changed mixed runs', () => {
  const source = content('<p data-source-paragraph="0:0" data-word-paragraph-font-features="3"><span data-word-font-features="3">Native</span></p>');
  const saved = content('<p data-source-paragraph="0:0"><b>Edited</b></p>'), original = structuredClone(saved);
  const migrated = hydrateWordFontFeatures(saved, source);
  expect(wordJSON(migrated.html).content![0].content![0].marks).toContainEqual(expect.objectContaining({ type: 'textStyle', attrs: expect.objectContaining({ wordFontFeatures: 3 }) }));
  expect(saved).toEqual(original);
  const mixed = content('<p data-source-paragraph="0:0"><span data-word-font-features="1">of</span><span data-word-font-features="0">fice</span></p>');
  expect(() => hydrateWordFontFeatures(saved, mixed)).toThrow(/missing mixed/);
  const same = hydrateWordFontFeatures(content('<p data-source-paragraph="0:0"><i>office</i></p>'), mixed);
  expect(wordJSON(same.html).content![0].content!.map((node) => node.marks!.find((m) => m.type === 'textStyle')!.attrs!.wordFontFeatures)).toEqual([1, 0]);
  expect(saved).toEqual(original);
});
