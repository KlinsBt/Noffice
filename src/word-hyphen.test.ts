import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { Editor, type JSONContent } from '@tiptap/core';
import JSZip from 'jszip';
import { readDocx } from './docx-import';
import { wordExtensions, wordJSON } from './word-extensions';
import { wordHyphenText } from './word-hyphen';
import { exportRetainedDocument } from './docx-preserve';
import { sanitizeWordContent, exportOffice } from './formats';
import { newFile } from './model';
import { contentFingerprint } from './office-preservation';
import { hydrateWordStructure, needsWordStructure } from './word-structure';
import { descendants, wordXml } from './word-xml';
import reference from '../tests/fixtures/word-soft-hyphens/reference.json';

vi.mock('mammoth', async original => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});
const bytes = (blob: Blob) => new Promise<ArrayBuffer>((resolve, reject) => {
  const reader = new FileReader();reader.onload = () => resolve(reader.result as ArrayBuffer);
  reader.onerror = () => reject(reader.error);reader.readAsArrayBuffer(blob);
});
const input = (sample: typeof reference.rows[number]) => {
  const data = Uint8Array.from(readFileSync(sample.path)).buffer;
  expect(createHash('sha256').update(new Uint8Array(data)).digest('hex')).toBe(sample.sourceHash);
  return data;
};
const nativeText = (doc: JSONContent) => doc.content!.map(paragraph =>
  (paragraph.content || []).map(run => run.type === 'wordHyphen'
    ? wordHyphenText(run.attrs?.kind, 'native') : run.text || '').join('') + '\r').join('');

for (const sample of reference.rows) it(`retains native character identities on import: ${sample.name}`, async () => {
  const content = sanitizeWordContent((await readDocx(input(sample))).content);
  expect(content.hyphenVersion).toBe(1);
  const json = wordJSON(content.html);
  expect(nativeText(json)).toBe(sample.native.sourceText);
  const hyphens = json.content!.flatMap(p => p.content || []).filter(node => node.type === 'wordHyphen');
  expect(hyphens).toHaveLength(22);
  expect(new Set(hyphens.map(node => node.attrs?.kind))).toEqual(new Set([sample.encoding === 'element' ? 'optional' : 'literal']));
});

for (const encoding of ['element', 'literal']) it(`edits ${encoding} controls with history and actual preserved DOCX`, async () => {
  const sample = reference.rows.find(row => row.name === `regular-160-${encoding}`)!;
  const data = input(sample), content = sanitizeWordContent((await readDocx(data)).content);
  const file = newFile('word', sample.name, content);
  file.original = { name: 'hyphens.docx', data };
  const editor = new Editor({ extensions: wordExtensions({ retainedEditRuns: true }), content: content.html });
  const sync = () => { content.html = editor.getHTML();content.initialEditSession = editor.state.doc.attrs.wordInitialEditSession || undefined; };
  try {
    sync();file.original.contentFingerprint = await contentFingerprint(content);
    expect(new Uint8Array(await bytes(await exportRetainedDocument(file)))).toEqual(new Uint8Array(data));
    const before = editor.getJSON();let at = -1;
    editor.state.doc.child(1).forEach((node, offset) => {
      if (at < 0 && node.type.name === 'wordHyphen') at = editor.state.doc.child(0).nodeSize + 1 + offset;
    });
    expect(at).toBeGreaterThan(0);editor.commands.deleteRange({ from: at, to: at + 1 });
    const after = editor.getJSON();expect(nativeText(after)).toBe(sample.native.editedText);
    editor.commands.undo();expect(editor.getJSON()).toEqual(before);
    editor.commands.redo();expect(editor.getJSON()).toEqual(after);sync();
    const output = await bytes(await exportRetainedDocument(file));
    const zip = await JSZip.loadAsync(output), original = await JSZip.loadAsync(data);
    for (const [path, entry] of Object.entries(original.files)) {
      if (entry.dir || ['word/document.xml', 'word/settings.xml'].includes(path)) continue;
      expect(await zip.file(path)!.async('uint8array')).toEqual(await entry.async('uint8array'));
    }
    const xml = wordXml(await zip.file('word/document.xml')!.async('string'));
    expect(descendants(xml, 'softHyphen')).toHaveLength(encoding === 'element' ? 21 : 0);
    expect(descendants(xml, 't').map(t => t.textContent).join('').split('\u00ad').length - 1).toBe(encoding === 'literal' ? 21 : 0);
    expect(nativeText(wordJSON((await readDocx(output)).content.html))).toBe(sample.native.editedText);
    expect(file.original.data).toBe(data);
  } finally { editor.destroy(); }
});

it('exports freshly authored adjacent optional and literal controls without changing their meanings', async () => {
  const file = newFile('word');if (file.content.kind !== 'word') throw Error();
  file.content.html = '<p>A<span data-word-hyphen="optional">\u00ad</span><span data-word-hyphen="optional">\u00ad</span>B<span data-word-hyphen="literal">-</span>C</p>';
  const zip = await JSZip.loadAsync(await bytes(await exportOffice(file)));
  const xml = wordXml(await zip.file('word/document.xml')!.async('string'));
  expect(descendants(xml, 'softHyphen')).toHaveLength(2);
  expect(descendants(xml, 't').map(t => t.textContent).join('')).toBe('AB\u00adC');
});

it.each(['unchanged', 'prefix', 'ambiguous'])('migrates legacy hyphen identity safely: %s', async mode => {
  const sample = reference.rows.find(row => row.name === 'regular-160-element')!, data = input(sample);
  const content = sanitizeWordContent((await readDocx(data, { legacyHyphens: true })).content);
  const file = newFile('word', 'Legacy hyphens', content);
  file.original = { name: 'hyphens.docx', data, contentFingerprint: await contentFingerprint(content) };
  if (mode === 'prefix') content.html = content.html.replace('antidisestab', 'Edited antidisestab');
  if (mode === 'ambiguous') content.html = content.html.replace('micro\u00adscope', 'new\u00adtext');
  expect(needsWordStructure(file)).toBe(true);
  const unchanged = structuredClone(file);
  if (mode === 'ambiguous') {
    await expect(hydrateWordStructure(file)).rejects.toThrow('ambiguous edited hyphen');
    expect(file).toEqual(unchanged);return;
  }
  const migrated = await hydrateWordStructure(file);
  expect(migrated.content.kind).toBe('word');if (migrated.content.kind !== 'word') throw Error();
  expect(needsWordStructure(migrated)).toBe(false);
  expect(nativeText(wordJSON(migrated.content.html))).toBe(mode === 'prefix'
    ? sample.native.sourceText.replace('antidisestab', 'Edited antidisestab') : sample.native.sourceText);
  expect(migrated.revision).toBe(file.revision);expect(migrated.original!.data).toBe(data);
  if (mode === 'unchanged') expect(new Uint8Array(await bytes(await exportRetainedDocument(migrated)))).toEqual(new Uint8Array(data));
  expect(file).toEqual(unchanged);
});

it('does not accept malformed hyphen markers as an inline character', () => {
  for (const html of ['<span data-word-hyphen="unknown">-</span>', '<span data-word-hyphen="literal">hidden text</span>'])
    expect(wordJSON(`<p>${html}</p>`).content![0].content!.every(node => node.type !== 'wordHyphen')).toBe(true);
});

it('inserts a hyphen with the active formatting and atomically restores it through history', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p><strong>AB</strong></p>' });
  try {
    editor.commands.setTextSelection(2); const before = editor.getJSON();
    expect(editor.commands.insertWordHyphen('optional')).toBe(true);
    const after = editor.getJSON(), control: JSONContent = after.content![0].content![1];
    expect(control.type).toBe('wordHyphen'); expect(control.attrs?.kind).toBe('optional');
    expect(control.marks?.some(mark => mark.type === 'bold')).toBe(true);
    editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
    editor.commands.redo(); expect(editor.getJSON()).toEqual(after);
    expect(editor.commands.insertWordHyphen('unknown' as 'optional')).toBe(false);
    expect(editor.getJSON()).toEqual(after);
  } finally { editor.destroy(); }
});

it.each(['attribute', 'content'])('preserves malformed source hyphens when paragraph deletion cannot map them: %s', async kind => {
  const sample = reference.rows.find(row => row.name === 'regular-160-element')!;
  const zip = await JSZip.loadAsync(input(sample));
  const document = wordXml(await zip.file('word/document.xml')!.async('string'));
  const control = descendants(document, 'softHyphen')[0];
  if (kind === 'attribute') control.setAttribute('unknown', 'preserve');
  else control.textContent = 'hidden';
  zip.file('word/document.xml', new XMLSerializer().serializeToString(document));
  const data = await zip.generateAsync({ type: 'arraybuffer' });
  const content = sanitizeWordContent((await readDocx(data)).content);
  const file = newFile('word', 'Malformed hyphen', content);
  file.original = { name: 'malformed.docx', data, contentFingerprint: await contentFingerprint(content) };
  const html = new DOMParser().parseFromString(content.html, 'text/html');
  html.querySelector('p')!.remove(); content.html = html.body.innerHTML;
  const unchanged = structuredClone(file);
  await expect(exportRetainedDocument(file)).rejects.toThrow('without losing document features');
  expect(file).toEqual(unchanged);
});
