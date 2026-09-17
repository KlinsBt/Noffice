import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { readDocxStructure } from './docx-sections';
import { wordXml } from './docx-import';
import { readWordStorySources, wordStoryPartTarget } from './word-story-source';
import { wordStoriesSchema } from './word-stories';
import reference from '../tests/fixtures/native-word-side-stories.json';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const main = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="h"/></w:sectPr></w:body></w:document>`;
const relationship = `<Relationship Id="h" Type="${R}/header" Target="header1.xml"/>`;
const rels = (body: string) => `<Relationships xmlns="${REL}">${body}</Relationships>`;
const header = `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`;
const structure = readDocxStructure(wordXml(main));
function packageWith(changes: Record<string, string | null> = {}) {
  const zip = new JSZip();
  const files: Record<string, string | null> = {
    'word/document.xml': main,
    'word/_rels/document.xml.rels': rels(relationship),
    'word/header1.xml': header,
    ...changes,
  };
  for (const [path, value] of Object.entries(files)) if (value !== null) zip.file(path, value);
  return zip;
}

function nativeText(node: Element): string {
  if (node.namespaceURI === W && node.localName === 't') return node.textContent || '';
  if (node.namespaceURI === W && node.localName === 'br') return '\v';
  return (
    [...node.children].map(nativeText).join('') +
    (node.namespaceURI === W && node.localName === 'p' ? '\r' : '')
  );
}

describe('independently authored native header/footer source identities', () => {
  for (const [name, sample] of Object.entries(reference.cases))
    it(`reads every native ${name} story without mixing it into body paragraphs`, async () => {
      const bytes = readFileSync(`tests/fixtures/${sample.sourceFile}`);
      expect(createHash('sha256').update(bytes).digest('hex')).toBe(sample.sourceSha256);
      const zip = await JSZip.loadAsync(bytes);
      const original = await zip.file('word/document.xml')!.async('string');
      const source = readDocxStructure(wordXml(original));
      const stories = await readWordStorySources(zip, source);
      expect(stories.evenAndOddHeaders).toBe(sample.native.sections[0].even);
      // Word retains referenced first/even parts even when those page options
      // are disabled. COM Exists describes their active story, not ZIP presence.
      expect(stories.parts).toHaveLength(
        source.sections[0].headers.length + source.sections[0].footers.length,
      );
      const slots = { default: 1, first: 2, even: 3 };
      for (const [key, kind] of [
        ['headers', 'Headers'],
        ['footers', 'Footers'],
      ] as const)
        for (const ref of source.sections[0][key]) {
          const part = stories.parts.find((p) => p.relationshipIds.includes(ref.relationshipId))!;
          expect(part).toBeDefined();
          const native = sample.native.sections[0].stories.find(
            (s) => s.kind === kind && s.slot === slots[ref.type as keyof typeof slots],
          )!;
          // This fixture only writes enabled stories; its inactive parts are
          // independently authored blank paragraphs and must also be retained.
          expect(nativeText(part.document.documentElement)).toBe(
            native.exists ? native.text : '\r',
          );
        }
      expect(await zip.file('word/document.xml')!.async('string')).toBe(original);
      const part = stories.parts[0],
        before = await zip.file(part.path)!.async('string');
      part.document.documentElement.textContent = 'Reading-copy annotation';
      expect(await zip.file(part.path)!.async('string')).toBe(before);
    });
});

it('keeps multiple relationship aliases attached to one source part', async () => {
  const doc = wordXml(
    main.replace('</w:sectPr>', '<w:headerReference w:type="first" r:id="h2"/></w:sectPr>'),
  );
  const zip = packageWith({
    'word/_rels/document.xml.rels': rels(relationship + relationship.replace('Id="h"', 'Id="h2"')),
  });
  const result = await readWordStorySources(zip, readDocxStructure(doc));
  expect(result.parts).toHaveLength(1);
  expect(result.parts[0].relationshipIds).toEqual(['h', 'h2']);
});

it.each([
  ['header1.xml', 'word/header1.xml'],
  ['/custom/header.xml', 'custom/header.xml'],
  ['../custom/header%20one.xml', 'custom/header one.xml'],
  ['./headers/../header1.xml', 'word/header1.xml'],
])('resolves internal relationship %s', (target, expected) => {
  expect(wordStoryPartTarget('word/document.xml', target)).toBe(expected);
});

it.each([
  '../../escape.xml',
  '%2e%2e/%2e%2e/escape.xml',
  'https://host/header.xml',
  '//host/header.xml',
  'header.xml?query',
  'header.xml#fragment',
  '..\\header.xml',
  'header%2fsecret.xml',
  '%00.xml',
  '%zz',
])('rejects unsafe or ambiguous target %s', (target) => {
  expect(() => wordStoryPartTarget('word/document.xml', target)).toThrow();
});

const brokenPackages: Record<string, string | null>[] = [
  { 'word/_rels/document.xml.rels': null },
  { 'word/_rels/document.xml.rels': rels('') },
  { 'word/_rels/document.xml.rels': rels(relationship + relationship) },
  { 'word/_rels/document.xml.rels': rels(relationship.replace('/>', ' TargetMode="External"/>')) },
  { 'word/header1.xml': null },
  { 'word/header1.xml': header.replaceAll('w:hdr', 'w:ftr') },
  { 'word/header1.xml': '<!DOCTYPE x [<!ENTITY y "bad">]>' + header },
  {
    'word/settings.xml': `<w:document xmlns:w="${W}"/>`,
    'word/_rels/document.xml.rels': rels(relationship + `<Relationship Id="settings" Type="${R}/settings" Target="settings.xml"/>`),
  },
];
it.each(brokenPackages)(
  'rejects broken referenced stories instead of losing their content %#',
  async (changes) => {
    await expect(readWordStorySources(packageWith(changes), structure)).rejects.toThrow();
  },
);

it('rejects ambiguous persisted identities and noncanonical part paths', () => {
  const part = {
    path: 'word/header1.xml',
    kind: 'header',
    relationshipIds: ['h'],
    html: '<p>Header</p>',
  };
  const value = { version: 1, evenAndOddHeaders: false, parts: [part] };
  expect(wordStoriesSchema.safeParse(value).success).toBe(true);
  expect(wordStoriesSchema.safeParse({ ...value, parts: [part, part] }).success).toBe(false);
  expect(
    wordStoriesSchema.safeParse({ ...value, parts: [part, { ...part, path: 'word/header2.xml' }] })
      .success,
  ).toBe(false);
  expect(
    wordStoriesSchema.safeParse({ ...value, parts: [{ ...part, path: '../../document.xml' }] })
      .success,
  ).toBe(false);
});
