import { it, expect, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { readDocx, wordXml, descendants, child, val } from './docx-import';
import { changedWordStoryCreation } from './word-story-create';
import { wordExtensions } from './word-extensions';
import { resolveWordSections } from './word-section-layout';
import { exportRetainedDocument } from './docx-preserve';
import { emptyStoryDocument } from './docx-empty-stories';
import { newFile } from './model';
import { sanitizeWordContent } from './formats';
import { contentFingerprint } from './office-preservation';
import { hydrateWordStructure, needsWordStructure } from './word-structure';
import reference from '../tests/fixtures/native-word-style-defaults.json';

vi.mock('mammoth', async (original) => {
  const actual = await original<typeof import('mammoth')>();
  return {
    ...actual,
    convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
      actual.convertToHtml(
        { buffer: Buffer.from(input.arrayBuffer) },
        options as Parameters<typeof actual.convertToHtml>[1],
      ),
  };
});
const bytes = (file: string) => Uint8Array.from(readFileSync('tests/fixtures/' + file)).buffer;
const blobBytes = (blob: Blob) =>
  new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });
function paragraph(html: string) {
  return new DOMParser().parseFromString(html, 'text/html').querySelector('p')!;
}
function expectNative(html: string, native: (typeof reference.cases)[0]['body']) {
  const p = paragraph(html);
  expect(p.style.fontFamily.replace(/^["']|["']$/g, '')).toBe(native.family);
  expect(p.style.fontSize).toBe(`${native.size}pt`);
  expect(Number.parseFloat(p.style.marginBottom) || 0).toBe(native.after);
  expect(Number(p.style.lineHeight)).toBeCloseTo(native.rule === 0 ? 1 : native.line / 12, 6);
  expect(p.textContent).toBe(native.text.replace(/\r$/, ''));
}
for (const sample of reference.cases)
  it(`uses native defaults for ${sample.mode} before and after exporting a created header`, async () => {
    const data = bytes(sample.file);
    const { content } = await readDocx(data);
    expectNative(content.html, sample.initial);
    const file = newFile('word', sample.mode, content);
    file.original = {
      name: sample.file,
      data,
      contentFingerprint: await contentFingerprint(content),
    };
    const editor = new Editor({
      extensions: wordExtensions(),
      content: content.stories!.emptyTemplates!.header!,
    });
    try {
      editor.commands.insertContent('Created header');
      content.stories = changedWordStoryCreation(
        content,
        resolveWordSections(content)[0].id,
        'header',
        'first',
        editor.getHTML(),
      );
    } finally {
      editor.destroy();
    }
    const exported = await blobBytes(await exportRetainedDocument(file));
    const restored = (await readDocx(exported)).content;
    expectNative(restored.html, sample.body);
    const id = resolveWordSections(restored)[0].headers.first!.relationshipId;
    expectNative(
      restored.stories!.parts.find((p) => p.relationshipIds.includes(id))!.html,
      sample.header,
    );
    const before = await JSZip.loadAsync(data),
      after = await JSZip.loadAsync(exported);
    // Body is unedited: synthesizing story styles must not materialize new body formatting.
    const bodyRuns = async (zip: JSZip) =>
      descendants(wordXml(await zip.file('word/document.xml')!.async('string')), 'r').map((r) =>
        new XMLSerializer().serializeToString(r),
      );
    expect(await bodyRuns(after)).toEqual(await bodyRuns(before));
    const rels = wordXml(await after.file('word/_rels/document.xml.rels')!.async('string'));
    expect(
      [...rels.documentElement.children].filter((r) => r.getAttribute('Type')?.endsWith('/styles')),
    ).toHaveLength(1);
    const types = wordXml(await after.file('[Content_Types].xml')!.async('string'));
    expect(
      [...types.documentElement.children].filter(
        (t) => t.getAttribute('PartName') === '/word/styles.xml',
      ),
    ).toHaveLength(1);
    const styles = wordXml(await after.file('word/styles.xml')!.async('string'));
    const normal = descendants(styles, 'style').find(
      (s) => val(s, 'type') === 'paragraph' && val(s, 'default') === '1',
    )!;
    expect(child(normal, 'rPr')).toBeUndefined();
  });

it('migrates unchanged legacy defaults, keeps original exports byte-identical and becomes idempotent', async () => {
  const sample = reference.cases[0],
    data = bytes(sample.file);
  const content = sanitizeWordContent(
    (await readDocx(data, { legacyStyleDefaults: true })).content,
  );
  content.stories!.templateVersion = 3;
  content.stories!.emptyTemplates = { header: null, footer: null };
  const file = newFile('word', 'Legacy defaults', content);
  file.original = {
    name: sample.file,
    data,
    contentFingerprint: await contentFingerprint(content),
  };
  const before = structuredClone(file);
  const restored = await hydrateWordStructure(file);
  if (restored.content.kind !== 'word') throw Error('Expected Word');
  expectNative(restored.content.html, sample.initial);
  expect(restored.content.styleDefaultsVersion).toBe(1);
  expect(restored.content.stories!.templateVersion).toBe(4);
  expect(needsWordStructure(restored)).toBe(false);
  expect(await hydrateWordStructure(restored)).toBe(restored);
  expect(restored.revision).toBe(file.revision);
  expect(new Uint8Array(await blobBytes(await exportRetainedDocument(restored)))).toEqual(
    new Uint8Array(data),
  );
  expect(file).toEqual(before);
});

it('rejects an ambiguous edited legacy default cascade without changing saved edits or originals', async () => {
  const sample = reference.cases[0],
    data = bytes(sample.file);
  const content = sanitizeWordContent(
    (await readDocx(data, { legacyStyleDefaults: true })).content,
  );
  const file = newFile('word', 'Edited legacy defaults', content);
  file.original = {
    name: sample.file,
    data,
    contentFingerprint: await contentFingerprint(content),
  };
  content.html = content.html.replace('Unformatted body', 'Saved edits at 10pt');
  const before = structuredClone(file);
  await expect(hydrateWordStructure(file)).rejects.toThrow('Export a Noffice backup');
  expect(file).toEqual(before);
});

it('upgrades unaffected edited snapshots without replacing their HTML or original identity', async () => {
  const data = bytes('word-story-base-missing-both.docx');
  const content = sanitizeWordContent(
    (await readDocx(data, { legacyStyleDefaults: true })).content,
  );
  const file = newFile('word', 'Unaffected legacy defaults', content);
  file.original = {
    name: 'source.docx',
    data,
    contentFingerprint: await contentFingerprint(content),
  };
  content.html = content.html.replace('B1-01', 'Saved explicit edit');
  const restored = await hydrateWordStructure(file);
  if (restored.content.kind !== 'word') throw Error('Expected Word');
  expect(restored.content.html).toBe(content.html);
  expect(restored.original).toBe(file.original);
  expect(restored.content.styleDefaultsVersion).toBe(1);
});

it('does not repoint conflicting styles relationships or partially write metadata on failure', async () => {
  const zip = await JSZip.loadAsync(bytes(reference.cases[0].file));
  const path = 'word/_rels/document.xml.rels';
  const rels = wordXml(await zip.file(path)!.async('string'));
  const r = rels.createElementNS(rels.documentElement.namespaceURI, 'Relationship');
  r.setAttribute('Id', 'other-styles');
  r.setAttribute(
    'Type',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  );
  r.setAttribute('Target', 'alternate-styles.xml');
  rels.documentElement.append(r);
  const xml = new XMLSerializer().serializeToString(rels);
  zip.file(path, xml);
  const types = await zip.file('[Content_Types].xml')!.async('string');
  await expect(emptyStoryDocument(zip, 'header')).rejects.toThrow('conflicting package metadata');
  expect(zip.file('word/styles.xml')).toBeNull();
  expect(await zip.file(path)!.async('string')).toBe(xml);
  expect(await zip.file('[Content_Types].xml')!.async('string')).toBe(types);
});
