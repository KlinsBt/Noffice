import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { readDocx, wordXml, child, val, descendants } from './docx-import';
import { wordExtensions, wordJSON } from './word-extensions';
import { exportOffice, importFile } from './formats';
import { newFile } from './model';
import { contentFingerprint } from './office-preservation';
import { hydrateWordStructure } from './word-structure';
import { wordParagraphSpaceInput, wordParagraphSpaceTwips } from './word-paragraph-spacing';
import { wordSelectionParagraphSpacing } from './word-paragraph-spacing-commands';
import native from '../tests/fixtures/native-word-paragraph-spacing.json';
import boundaries from '../tests/fixtures/native-word-paragraph-spacing-boundaries.json';

vi.mock('mammoth', async original => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});
const fixture = (name: string) => new Uint8Array(readFileSync(`tests/fixtures/word-paragraph-spacing-${name}.docx`)).buffer;
const load = async (name: string) => {
  const data = fixture(name);
  return importFile(Object.assign(new File([data], name + '.docx'), { arrayBuffer: async () => data }));
};
const blobBytes = (blob: Blob) => new Promise<ArrayBuffer>((resolve, reject) => {
  const reader = new FileReader(); reader.onload = () => resolve(reader.result as ArrayBuffer); reader.onerror = reject; reader.readAsArrayBuffer(blob);
});

it('matches every measured native point/line rounding and rejection boundary', () => {
  for (const row of boundaries.rows) {
    const lines = row.property.startsWith('LineUnit');
    const space = wordParagraphSpaceInput(lines ? 'lines' : 'points', row.value);
    expect(!!space, JSON.stringify(row)).toBe(row.accepted);
    if (!space) continue;
    const side = row.property.endsWith('After') ? 'after' : 'before';
    const record = row as Record<string, unknown>;
    expect(wordParagraphSpaceTwips(space) / 20, JSON.stringify(row)).toBeCloseTo(Number(record[side]), 3);
    if (lines && space.unit === 'lines') expect(space.value / 100).toBeCloseTo(Number(record[side + 'Lines']), 3);
  }
  for (const invalid of ['', ' ', NaN, Infinity, -Infinity, 'bad', {}, null])
    expect(wordParagraphSpaceInput('points', invalid)).toBeNull();
});

for (const row of native.rows) it(`imports native spacing precedence and preserves a text edit: ${row.name}`, async () => {
  const file = await load(row.name);
  if (file.content.kind !== 'word') throw Error('Word fixture');
  const json = wordJSON(file.content.html), p = json.content![2], attrs = p.attrs!;
  const before = row.geometry.advances[1] - 40, after = row.geometry.advances[2] - 40;
  if (!row.name.startsWith('contextual')) {
    expect(Math.abs(parseFloat(attrs.spaceBefore || '0') - before)).toBeLessThanOrEqual(.15);
    expect(Math.abs(parseFloat(attrs.spaceAfter || '0') - after)).toBeLessThanOrEqual(.15);
  } else expect(attrs.paragraphContextualSpacing).toBe(true);
  const editor = new Editor({ extensions: wordExtensions(), content: file.content.html });
  try {
    editor.commands.setTextSelection(1); editor.commands.insertContent('Edited');
    file.content.html = editor.getHTML();
    const exported = await blobBytes(await exportOffice(file));
    const actual = (await readDocx(exported)).content;
    const result = wordJSON(actual.html).content![2].attrs!;
    for (const key of ['spaceBefore', 'spaceAfter', 'paragraphSpaceBefore', 'paragraphSpaceAfter', 'paragraphContextualSpacing'])
      expect(result[key], key).toEqual(attrs[key]);
    const originalZip = await JSZip.loadAsync(fixture(row.name)), outputZip = await JSZip.loadAsync(exported);
    for (const name of ['word/styles.xml', 'word/settings.xml'])
      expect(await outputZip.file(name)!.async('string')).toBe(await originalZip.file(name)!.async('string'));
  } finally { editor.destroy(); }
});

it('switches inherited modes to points in one undo event and exports explicit overrides', async () => {
  const file = await load('line-inherited-point-override');
  if (file.content.kind !== 'word') throw Error('Word fixture');
  const editor = new Editor({ extensions: wordExtensions(), content: file.content.html });
  try {
    editor.commands.setTextSelection(7);
    const original = editor.getJSON();
    expect(wordSelectionParagraphSpacing(editor.state).before).toEqual({ unit: 'lines', value: 100 });
    expect(editor.can().setWordParagraphSpacing({ before: { unit: 'points', value: 240 } })).toBe(true);
    expect(editor.can().undo()).toBe(false);
    editor.commands.setWordParagraphSpacing({ before: { unit: 'points', value: 240 } });
    const changed = editor.getJSON(); expect(changed).not.toEqual(original);
    editor.commands.setWordParagraphSpacing({ before: { unit: 'points', value: 240 } });
    editor.commands.undo(); expect(editor.getJSON()).toEqual(original); expect(editor.can().undo()).toBe(false);
    editor.commands.redo(); expect(editor.getJSON()).toEqual(changed);
    file.content.html = editor.getHTML();
    const zip = await JSZip.loadAsync(await exportOffice(file));
    const doc = wordXml(await zip.file('word/document.xml')!.async('string'));
    const spacing = child(child(descendants(doc, 'p')[2], 'pPr')!, 'spacing');
    expect(val(spacing, 'before')).toBe('240'); expect(val(spacing, 'beforeLines')).toBe('0'); expect(val(spacing, 'beforeAutospacing')).toBe('0');
    expect(wordJSON((await readDocx(await zip.generateAsync({ type: 'arraybuffer' }))).content.html).content![2].attrs!.paragraphSpaceBefore).toBeNull();
  } finally { editor.destroy(); }
});

it('shows mixed semantic modes and applies only the changed field atomically', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p style="margin-top:12pt;margin-bottom:6pt">A</p><p data-word-space-before="{&quot;unit&quot;:&quot;lines&quot;,&quot;value&quot;:100}" style="margin-top:12pt;margin-bottom:8pt">B</p>' });
  try {
    editor.commands.selectAll(); const original = editor.getJSON();
    expect(wordSelectionParagraphSpacing(editor.state).before).toBeNull();
    expect(editor.commands.setWordParagraphSpacing({ after: { unit: 'points', value: -1 } })).toBe(false);
    expect(editor.getJSON()).toEqual(original); expect(editor.can().undo()).toBe(false);
    editor.commands.setWordParagraphSpacing({ after: { unit: 'points', value: 200 } });
    expect(editor.getJSON().content!.map(p => p.attrs!.paragraphSpaceBefore)).toEqual([null, { unit: 'lines', value: 100 }]);
    editor.commands.undo(); expect(editor.getJSON()).toEqual(original);
  } finally { editor.destroy(); }
});

it('keeps contextual suppression out of saved semantics and updates it after a neighbor style change', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p data-word-paragraph-style="Normal">A</p><p data-word-paragraph-style="Normal" data-word-contextual-spacing="true" style="margin-top:8pt;margin-bottom:12pt">B</p><p data-word-paragraph-style="Other">C</p>' });
  try {
    const p = () => editor.view.dom.querySelectorAll('p')[1];
    expect(p().classList.contains('word-contextual-before')).toBe(true);
    expect(p().style.marginTop).toBe('8pt'); expect(p().style.marginBottom).toBe('12pt');
    expect(editor.getHTML()).toContain('margin-top: 8pt');
    editor.view.dispatch(editor.state.tr.setNodeAttribute(0, 'paragraphStyleId', 'Other'));
    expect(p().style.marginTop).toBe('8pt'); expect(p().classList.contains('word-contextual-before')).toBe(false);
    editor.commands.undo(); expect(p().classList.contains('word-contextual-before')).toBe(true);
  } finally { editor.destroy(); }
});

it('writes fresh line/auto modes and contextual spacing as actual OOXML', async () => {
  const file = newFile('word', 'Spacing');
  if (file.content.kind !== 'word') throw Error('Word fixture');
  file.content.html = '<p data-word-space-before="{&quot;unit&quot;:&quot;lines&quot;,&quot;value&quot;:50}" data-word-space-after="{&quot;unit&quot;:&quot;auto&quot;}" data-word-contextual-spacing="true" style="margin-top:6pt;margin-bottom:14pt">Fresh</p>';
  const zip = await JSZip.loadAsync(await exportOffice(file));
  const xml = wordXml(await zip.file('word/document.xml')!.async('string'));
  expect(descendants(xml, 'spacing')).toHaveLength(1);
  const order = [...descendants(xml, 'pPr')[0].children].map(p => p.localName);
  expect(order.indexOf('spacing')).toBeLessThan(order.indexOf('contextualSpacing'));
  const result = wordJSON((await readDocx(await zip.generateAsync({ type: 'arraybuffer' }))).content.html).content![0].attrs!;
  expect(result.paragraphSpaceBefore).toEqual({ unit: 'lines', value: 50 });
  expect(result.paragraphSpaceAfter).toEqual({ unit: 'auto' }); expect(result.paragraphContextualSpacing).toBe(true);
});

it('migrates retained missing modes while preserving edited points, revisions and original bytes', async () => {
  const file = await load('line-before-no-fallback');
  if (file.content.kind !== 'word') throw Error('Word fixture');
  file.content = (await readDocx(fixture('line-before-no-fallback'), { legacyParagraphSpacing: true })).content;
  file.original!.contentFingerprint = await contentFingerprint(file.content);
  const before = structuredClone(file), migrated = await hydrateWordStructure(file);
  expect(file).toEqual(before); expect(migrated.revision).toBe(file.revision);
  expect(Buffer.from(await blobBytes(await exportOffice(migrated)))).toEqual(Buffer.from(fixture('line-before-no-fallback')));
  const dom = new DOMParser().parseFromString(file.content.html, 'text/html');
  dom.querySelectorAll('p')[2].style.marginTop = '8pt'; dom.querySelectorAll('p')[0].textContent = 'Edited';
  file.content.html = dom.body.innerHTML;
  const edited = await hydrateWordStructure(file);
  if (edited.content.kind !== 'word') throw Error('Word fixture');
  const p = wordJSON(edited.content.html).content![2];
  expect(p.attrs!.spaceBefore).toBe('8pt'); expect(p.attrs!.paragraphSpaceBefore).toBeNull();
  const result = (await readDocx(await blobBytes(await exportOffice(edited)))).content;
  expect(wordJSON(result.html).content![2].attrs!.spaceBefore).toBe('8pt');
  file.content.html += '<p>Ambiguous split</p>';
  await expect(hydrateWordStructure(file)).rejects.toThrow(/backup/);
});
