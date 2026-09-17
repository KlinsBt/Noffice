import { expect, it, vi } from 'vitest';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import native from '../tests/fixtures/native-word-compatibility.json';
import { readWordSettings } from './docx-settings';
import { readWordCompatibility, wordCompatibilitySchema } from './word-compatibility';
import { readDocx, wordXml, WORD_NS } from './docx-import';
import { contentSchema, newFile } from './model';
import { hydrateWordStructure, needsWordStructure } from './word-structure';
import { contentFingerprint } from './office-preservation';
import { exportRetainedDocument } from './docx-preserve';
import { writeWordStoryOptions } from './docx-story-options';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';

vi.mock('mammoth', async (original) => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});
const relationships = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const relation = (target = 'settings.xml', extra = '') => `<Relationship Id="settings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="${target}" ${extra}/>`;
const packageWith = (body: string, target = 'settings.xml') => new JSZip()
  .file('word/_rels/document.xml.rels', relationships(relation(target)))
  .file('word/settings.xml', `<w:settings xmlns:w="${WORD_NS}">${body}</w:settings>`);
const data = async (name: string) => Uint8Array.from(await readFile(`tests/fixtures/word-compatibility-${name}.docx`)).buffer;
const blobBytes = (blob: Blob) => new Promise<ArrayBuffer>((resolve, reject) => {
  const reader = new FileReader();reader.onload = () => resolve(reader.result as ArrayBuffer);
  reader.onerror = () => reject(reader.error);reader.readAsArrayBuffer(blob);
});
const unresolved = new Set(['mode-16', 'mode-0', 'mode-invalid', 'duplicate-forward', 'duplicate-reverse']);

for (const row of native.rows) it(`retains measured native compatibility provenance: ${row.name}`, async () => {
  const bytes = await data(row.name);
  expect(createHash('sha256').update(Buffer.from(bytes)).digest('hex')).toBe(row.inputHash);
  const settings = await readWordSettings(await JSZip.loadAsync(bytes));
  const state = readWordCompatibility(settings);
  expect(wordCompatibilitySchema.parse(JSON.parse(JSON.stringify(state)))).toEqual(state);
  expect(state.settingsPath).toBe(row.name === 'renamed-mode15' ? 'word/metadata/settings-custom.xml' :
    ['settings-missing', 'orphan-settings'].includes(row.name) ? null : 'word/settings.xml');
  if (unresolved.has(row.name)) {
    expect(state).toMatchObject({ mode: null, modeOrigin: 'unresolved', wordPerfectJustification: null });
  } else {
    expect(state.mode).toBe(row.mode);
    expect(state.wordPerfectJustification).toBe(row.wpJustification);
  }
  const imported = await readDocx(bytes);
  expect(imported.content.docxStructure?.compatibility).toEqual(state);
  expect(imported.messages.some((message) => message.message.includes('unresolved Word compatibility'))).toBe(unresolved.has(row.name));
});

it.each([
  ['external', relation('https://example.invalid/settings.xml', 'TargetMode="External"')],
  ['duplicate settings', relation() + relation('other.xml')],
  ['duplicate ID', relation() + '<Relationship Id="settings" Type="unrelated" Target="other.xml"/>'],
  ['missing ID', relation().replace('Id="settings"', '')],
  ['escape', relation('../../settings.xml')],
  ['encoded slash', relation('metadata%2fsettings.xml')],
  ['missing part', relation('missing.xml')],
])('rejects %s settings relationships without falling back to an orphan file', async (_, body) => {
  const zip = packageWith('');
  zip.file('word/_rels/document.xml.rels', relationships(body));
  await expect(readWordSettings(zip)).rejects.toThrow();
  expect(await zip.file('word/settings.xml')!.async('string')).toBe(`<w:settings xmlns:w="${WORD_NS}"></w:settings>`);
});

it('ignores unreachable settings and wrong-namespace relationships', async () => {
  const zip = packageWith('');
  zip.file('word/_rels/document.xml.rels', relationships(relation().replace('<Relationship ', '<Relationship xmlns="urn:wrong" ')));
  expect(await readWordSettings(zip)).toBeUndefined();
  zip.remove('word/_rels/document.xml.rels');
  expect(readWordCompatibility(await readWordSettings(zip))).toMatchObject({ mode: 12, settingsPath: null });
});

it.each(['', '<w:wrong xmlns:w="'+WORD_NS+'"/>', '<settings/>', '<!DOCTYPE x [<!ENTITY a "x">]><x>&a;</x>'])('rejects malformed settings %s', async (text) => {
  const zip = packageWith('');zip.file('word/settings.xml', text);
  await expect(readWordSettings(zip)).rejects.toThrow();
});

it('does not read nested/foreign compatibility settings and bounds repeated source values', async () => {
  const mode = '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>';
  let zip = packageWith(`<w:other><w:compat>${mode}</w:compat></w:other>`);
  expect(readWordCompatibility(await readWordSettings(zip)).mode).toBe(12);
  zip = packageWith(`<w:compat>${mode}</w:compat><w:compat>${mode}</w:compat>`);
  expect(readWordCompatibility(await readWordSettings(zip)).mode).toBeNull();
  expect(() => readWordCompatibility({ path: 'word/settings.xml', document: wordXml(`<w:settings xmlns:w="${WORD_NS}"><w:compat>${mode.repeat(33)}</w:compat></w:settings>`) })).toThrow('excessive');
});

it.each(['mode-14', 'renamed-mode15', 'orphan-settings'])('hydrates missing %s provenance while preserving distinct edits, original identity and revision', async (name) => {
  const bytes = await data(name), imported = (await readDocx(bytes)).content;
  const file = newFile('word', name, imported);
  if (file.content.kind !== 'word') throw Error();
  file.original = { name: name+'.docx', data: bytes, contentFingerprint: await contentFingerprint(file.content) };
  file.content.html = file.content.html.replace('Body1', 'Saved body edit');
  file.content.stories!.parts[0].html = file.content.stories!.parts[0].html.replace('Alpha', 'Saved Alpha');
  delete file.content.docxStructure!.compatibility;
  const before = structuredClone(file), fingerprint = await contentFingerprint(file.content);
  expect(needsWordStructure(file)).toBe(true);
  const hydrated = await hydrateWordStructure(file);
  expect(hydrated.content).toMatchObject({ html: file.content.html, stories: file.content.stories });
  expect(hydrated.original).toBe(file.original);expect(hydrated.revision).toBe(file.revision);
  expect(await contentFingerprint(hydrated.content)).toBe(fingerprint);
  expect(file).toEqual(before);expect(needsWordStructure(hydrated)).toBe(false);
  expect(await hydrateWordStructure(hydrated)).toBe(hydrated);
  expect(contentSchema.parse(JSON.parse(JSON.stringify(hydrated.content)))).toEqual(hydrated.content);
});

it('rejects broken source settings during migration without mutating saved work', async () => {
  const bytes = await data('renamed-mode15'), content = (await readDocx(bytes)).content;
  const file = newFile('word', 'Saved work', content);if (file.content.kind !== 'word') throw Error();
  delete file.content.docxStructure!.compatibility;file.content.html = file.content.html.replace('Body1','Saved body edit');
  const zip = await JSZip.loadAsync(bytes);zip.remove('word/metadata/settings-custom.xml');
  file.original = { name: 'broken.docx', data: await zip.generateAsync({ type: 'arraybuffer' }) };
  const before = structuredClone(file);
  await expect(hydrateWordStructure(file)).rejects.toThrow('settings are missing');
  expect(file).toEqual(before);
});

it('edits a renamed-settings package and preserves compatibility, unrelated parts and the relationship', async () => {
  const bytes = await data('renamed-mode15'), content = (await readDocx(bytes)).content;
  const file = newFile('word', 'Renamed settings', content);if (file.content.kind !== 'word') throw Error();
  file.original = { name: 'renamed.docx', data: bytes, contentFingerprint: await contentFingerprint(content) };
  // Backup validation canonicalizes object key order; that is not a source edit.
  file.content = contentSchema.parse(file.content);
  if (file.content.kind !== 'word') throw Error();
  file.content.html = file.content.html.replace('Body1', 'Edited body');
  const result = await blobBytes(await exportRetainedDocument(file)), zip = await JSZip.loadAsync(result);
  expect(zip.file('word/settings.xml')).toBeNull();
  expect(readWordCompatibility(await readWordSettings(zip))).toEqual(content.docxStructure!.compatibility);
  const original = await JSZip.loadAsync(bytes);
  for (const entry of Object.values(original.files)) if (!entry.dir && !['word/document.xml', 'word/metadata/settings-custom.xml'].includes(entry.name))
    expect(await zip.file(entry.name)!.async('uint8array'), entry.name).toEqual(await entry.async('uint8array'));
  expect((await readDocx(result!)).content.html).toContain('Edited body');
  file.content.docxStructure!.compatibility!.mode = 14;
  await expect(exportRetainedDocument(file)).rejects.toThrow('source-section identities');
});

it('writes odd/even options to the referenced settings while preserving compatibility', async () => {
  const bytes = await data('renamed-mode15'), original = (await readDocx(bytes)).content, current = structuredClone(original);
  current.stories!.evenAndOddHeaders = !original.stories!.evenAndOddHeaders;
  const zip = await JSZip.loadAsync(bytes), main = wordXml(await zip.file('word/document.xml')!.async('string'));
  const before = readWordCompatibility(await readWordSettings(zip));
  expect(await writeWordStoryOptions(zip, main.getElementsByTagNameNS(WORD_NS,'body')[0], original, current)).toEqual({ bodyChanged: false, settingsChanged: true });
  expect(zip.file('word/settings.xml')).toBeNull();
  expect(readWordCompatibility(await readWordSettings(zip))).toEqual(before);
  expect((await readDocx(await zip.generateAsync({ type: 'arraybuffer' }))).content.stories!.evenAndOddHeaders).toBe(current.stories!.evenAndOddHeaders);
});

it.each(['mode-15', 'renamed-mode15'])('preserves odd/even options alongside distinct body and header typing sessions: %s', async (name) => {
  const bytes = await data(name), content = (await readDocx(bytes)).content;
  const file = newFile('word', name, content);if(file.content.kind !== 'word') throw Error();
  file.original = { name: name+'.docx', data: bytes, contentFingerprint: await contentFingerprint(content) };
  const story = file.content.stories!.parts.find((part) => part.html.includes('Alpha'))!;
  const bodyEditor = new Editor({ extensions: wordExtensions({ retainedEditRuns: true }), content: file.content.html });
  const storyEditor = new Editor({ extensions: wordExtensions({ retainedEditRuns: true }), content: story.html });
  try {
    bodyEditor.commands.setTextSelection(bodyEditor.state.doc.firstChild!.nodeSize-1);bodyEditor.commands.insertContent(' edited');
    storyEditor.commands.setTextSelection(1);storyEditor.commands.insertContent('Saved ');
    file.content.html=bodyEditor.getHTML();story.html=storyEditor.getHTML();file.content.stories!.evenAndOddHeaders=true;
    const output = await blobBytes(await exportRetainedDocument(file));
    const imported = (await readDocx(output)).content;
    expect(imported.stories!.evenAndOddHeaders).toBe(true);
    expect(new DOMParser().parseFromString(imported.html,'text/html').querySelector('p')!.textContent).toBe('Body1 edited');
    expect(imported.stories!.parts.some((part)=>part.html.includes('Saved '))).toBe(true);
    expect(imported.docxStructure!.compatibility).toEqual(content.docxStructure!.compatibility);
  } finally { bodyEditor.destroy();storyEditor.destroy(); }
});
