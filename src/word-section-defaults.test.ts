import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { readDocx, child, descendants, val, wordXml } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { sanitizeWordContent } from './formats';
import { newFile } from './model';
import { contentFingerprint } from './office-preservation';
import { resolveWordSections } from './word-section-layout';
import { wordFinalSectionProperties } from './word-section-defaults';
import { hydrateWordStructure, needsWordStructure } from './word-structure';
import { WordStoryState, initializeWordStories } from './word-story-edit';
import { WordEditorSections, updateWordSectionSource } from './word-editor-sections';
import { wordExtensions } from './word-extensions';
import reference from '../tests/fixtures/native-word-implicit-final-sections.json';

vi.mock('mammoth', async (original) => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});
const bytes = (blob: Blob) => new Promise<ArrayBuffer>((resolve, reject) => {
  const reader = new FileReader(); reader.onload = () => resolve(reader.result as ArrayBuffer);
  reader.onerror = () => reject(reader.error); reader.readAsArrayBuffer(blob);
});
const input = (sample: typeof reference.rows[number]) => {
  const data = Uint8Array.from(readFileSync(sample.fixture)).buffer;
  expect(createHash('sha256').update(new Uint8Array(data)).digest('hex')).toBe(sample.sourceHash);
  return data;
};

for (const sample of reference.rows) {
  it(`resolves independently measured ${sample.name} geometry without changing retained XML`, async () => {
    const content = sanitizeWordContent((await readDocx(input(sample))).content);
    const native = sample.states.find(s => s.stage === 'source')!.sourceSnapshot.sections;
    const sections = resolveWordSections(content);
    expect(sections).toHaveLength(native.length);
    sections.forEach((section, i) => {
      for (const key of ['width', 'height'] as const)
        expect(section[key]).toBe(Math.round(native[i][key] * 20));
      for (const key of ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const)
        expect(section.margins[key]).toBe(Math.round(native[i][key] * 20));
      expect(section.differentFirstPage).toBe(!!native[i].first);
    });
    expect(content.paper).toBe(sample.final === 'omitted' ? 'a4' : 'letter');
    expect(content.orientation).toBe('portrait');
    if (sample.final === 'omitted') expect(content.docxStructure!.sections.at(-1)!.propertiesXml).toBeNull();
    else expect(wordXml(content.docxStructure!.sections.at(-1)!.propertiesXml!).documentElement.children).toHaveLength(0);
  });

  for (const action of sample.sections > 1 ? ['insert', 'delete-break'] : ['insert'])
    it(`preserves ${sample.name} ${action}, history and actual section/story package contents`, async () => {
      const data = input(sample), content = sanitizeWordContent((await readDocx(data)).content);
      const file = newFile('word', sample.name, content);
      file.original = { name: sample.name + '.docx', data };
      const editor = new Editor({ extensions: [...wordExtensions(), WordStoryState,
        WordEditorSections.configure({ source: content.docxStructure })],
        content: content.html, parseOptions: { preserveWhitespace: true } });
      try {
        updateWordSectionSource(editor, content.docxStructure); initializeWordStories(editor, content);
        const sync = () => {
          content.html = editor.getHTML(); content.stories = editor.state.doc.attrs.wordStories;
          content.sectionState = editor.state.doc.attrs.wordSectionState;
        };
        sync(); file.original.contentFingerprint = await contentFingerprint(content);
        expect(new Uint8Array(await bytes(await exportRetainedDocument(file)))).toEqual(new Uint8Array(data));
        const before = editor.state.doc.toJSON();
        if (action === 'insert') editor.commands.insertContentAt(1, 'Edited ');
        else {
          let boundary = 0;
          for (let i = 0; i < 2 * (sample.sections - 1); i++) boundary += editor.state.doc.child(i).nodeSize;
          editor.commands.setTextSelection({ from: boundary - 1, to: boundary + 1 });
          editor.view.dispatch(closeHistory(editor.state.tr));
          expect(editor.commands.deleteWordBoundary()).toBe(true);
          editor.view.dispatch(closeHistory(editor.state.tr));
        }
        const after = editor.state.doc.toJSON(); editor.commands.undo();
        expect(editor.state.doc.toJSON()).toEqual(before); editor.commands.redo();
        expect(editor.state.doc.toJSON()).toEqual(after); sync();
        const output = await bytes(await exportRetainedDocument(file));
        const reopened = sanitizeWordContent((await readDocx(output)).content);
        const expected = sample.states.find(s => s.stage === action)!.savedSnapshot;
        expect(resolveWordSections(reopened)).toHaveLength(expected.sections.length);
        const originalZip = await JSZip.loadAsync(data), exportedZip = await JSZip.loadAsync(output);
        for (const name of Object.keys(originalZip.files).filter(n => n.startsWith('word/header') || n.startsWith('word/footer')))
          expect(await exportedZip.file(name)!.async('string')).toBe(await originalZip.file(name)!.async('string'));
        const xml = wordXml(await exportedZip.file('word/document.xml')!.async('string'));
        const final = child(descendants(xml, 'body')[0], 'sectPr')!;
        expect(val(child(final, 'pgSz'), 'w')).toBe(String(Math.round(expected.sections.at(-1)!.width * 20)));
        expect(val(child(final, 'pgSz'), 'h')).toBe(String(Math.round(expected.sections.at(-1)!.height * 20)));
        expect(descendants(xml, 't').map(t => t.textContent).join('')).toBe(expected.text.replace(/[\r\f]/g, ''));
        expect(file.original.data).toBe(data);
      } finally { editor.destroy(); }
    });
}

for (const explicit of [false, true]) it(`migrates old final-section controls while retaining edits and explicit choices=${explicit}`, async () => {
  const sample = reference.rows.find(r => r.name === '3-sections-omitted')!, data = input(sample);
  const content = sanitizeWordContent((await readDocx(data, { legacySectionDefaults: true })).content);
  expect(content.orientation).toBe('landscape');
  const file = newFile('word', 'Legacy sections', content);
  file.original = { name: 'legacy.docx', data, contentFingerprint: await contentFingerprint(content) };
  if (explicit) content.pageOverrides = { orientation: true, paper: true };
  content.html = content.html.replace('first paragraph.', 'edited paragraph.');
  const migrated = await hydrateWordStructure(file);
  expect(migrated.content.kind).toBe('word');
  if (migrated.content.kind !== 'word') throw Error();
  expect(migrated.content.html).toContain('edited paragraph.');
  expect(migrated.content.orientation).toBe(explicit ? 'landscape' : 'portrait');
  expect(migrated.content.pageOverrides).toEqual(content.pageOverrides);
  expect(migrated.original).toBe(file.original);
  expect(migrated.revision).toBe(file.revision);
  expect(needsWordStructure(migrated)).toBe(false);
  expect(await hydrateWordStructure(migrated)).toBe(migrated);
});

it('rebinds an unchanged legacy paper baseline without a storage revision or original-byte change', async () => {
  const sample = reference.rows.find(r => r.name === '2-sections-omitted')!, data = input(sample);
  const content = sanitizeWordContent((await readDocx(data, { legacySectionDefaults: true })).content);
  expect(content.paper).toBe('letter');
  const file = newFile('word', 'Unchanged legacy', content);
  file.original = { name: 'legacy.docx', data, contentFingerprint: await contentFingerprint(content) };
  const migrated = await hydrateWordStructure(file);
  expect(migrated.original!.contentFingerprint).toBe(await contentFingerprint(migrated.content));
  expect(migrated.original!.data).toBe(data); expect(migrated.revision).toBe(0);
  expect(new Uint8Array(await bytes(await exportRetainedDocument(migrated)))).toEqual(new Uint8Array(data));
});

it('leaves partial properties and unqualified compatibility modes on their existing path', () => {
  const partial = wordXml('<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgSz w:w="12240"/></w:sectPr>').documentElement;
  expect(wordFinalSectionProperties(partial, 15)).toBe(partial);
  for (const mode of [undefined, null, 14, 16]) expect(wordFinalSectionProperties(null, mode)).toBeNull();
});
