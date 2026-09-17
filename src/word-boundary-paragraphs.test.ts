import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { readDocx, child, descendants, val, wordXml } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { newFile } from './model';
import { contentFingerprint } from './office-preservation';
import { wordExtensions } from './word-extensions';
import { WordEditorSections, updateWordSectionSource } from './word-editor-sections';
import reference from '../tests/fixtures/word-boundary-delete/paragraphs.json';

vi.mock('mammoth', async original => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});
const bytes = (blob: Blob) => new Promise<ArrayBuffer>((resolve, reject) => {
  const reader = new FileReader(); reader.onload = () => resolve(reader.result as ArrayBuffer);
  reader.onerror = () => reject(reader.error); reader.readAsArrayBuffer(blob);
});

for (const sample of reference.rows) it(`preserves native paragraph layout and following mark through ${sample.name} export`, async () => {
  const data = Uint8Array.from(readFileSync(sample.fixture)).buffer;
  expect(createHash('sha256').update(new Uint8Array(data)).digest('hex')).toBe(sample.sourceHash);
  for (const native of sample.states) {
    const { content } = await readDocx(data), file = newFile('word', sample.name, content);
    file.original = { name: sample.name + '.docx', data, contentFingerprint: await contentFingerprint(content) };
    const editor = new Editor({ extensions: [...wordExtensions(), WordEditorSections.configure({ source: content.docxStructure })],
      content: content.html, parseOptions: { preserveWhitespace: true } });
    try {
      updateWordSectionSource(editor, content.docxStructure);
      const from = editor.state.doc.firstChild!.nodeSize - 1;
      editor.commands.setTextSelection({ from, to: from + 2 }); const before = editor.getJSON();
      expect(editor.commands.deleteWordBoundary(native.smart)).toBe(true);
      const after = editor.getJSON(); editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
      editor.commands.redo(); expect(editor.getJSON()).toEqual(after);
      content.html = editor.getHTML(); content.sectionState = editor.state.doc.attrs.wordSectionState;
      const output = await bytes(await exportRetainedDocument(file));
      const imported = await readDocx(output);
      const reloaded = new Editor({ extensions: wordExtensions(), content: imported.content.html });
      try {
        const attrs = reloaded.state.doc.firstChild!.attrs;
        expect(attrs.paragraphFontFamily?.replace(/^"(.*)"$/, '$1')).toBe(native.paragraphProperties.markFamily);
        expect(attrs.paragraphFontSize).toBe(native.paragraphProperties.markSize + 'pt');
      } finally { reloaded.destroy(); }
      const zip = await JSZip.loadAsync(output), xml = wordXml(await zip.file('word/document.xml')!.async('string'));
      const pr = child(descendants(xml, 'body')[0].children[0], 'pPr')!;
      const spacing = child(pr, 'spacing'), indent = child(pr, 'ind'), mark = child(pr, 'rPr');
      const on = (element: Element | null | undefined) => !!element && !['0', 'false', 'off'].includes(val(element));
      expect(Number(val(spacing, 'line')) / 20).toBe(native.paragraphProperties.line);
      expect(val(spacing, 'lineRule')).toBe('exact');
      expect(Number(val(spacing, 'before')) / 20).toBe(native.paragraphProperties.before);
      expect(Number(val(spacing, 'after')) / 20).toBe(native.paragraphProperties.after);
      expect(Number(val(indent, 'left') || val(indent, 'start')) / 20).toBe(native.paragraphProperties.left);
      expect(Number(val(indent, 'right') || val(indent, 'end')) / 20).toBe(native.paragraphProperties.right);
      expect(Number(val(indent, 'firstLine')) / 20).toBe(native.paragraphProperties.firstLine);
      expect(['left', 'center', 'right', 'both'].indexOf(val(child(pr, 'jc')) || 'left')).toBe(native.paragraphProperties.alignment);
      expect(on(child(pr, 'keepNext'))).toBe(!!native.paragraphProperties.keepNext);
      expect(on(child(pr, 'keepLines'))).toBe(!!native.paragraphProperties.keepLines);
      expect(on(mark && child(mark, 'b'))).toBe(!!native.paragraphProperties.markBold);
      expect(on(mark && child(mark, 'i'))).toBe(!!native.paragraphProperties.markItalic);
      expect(file.original.data).toBe(data);
    } finally { editor.destroy(); }
  }
});

for (const crossPart of [false, true]) it(`rejects ${crossPart ? 'cross-part' : 'unknown'} mark provenance without changing originals`, async () => {
  const path = crossPart ? 'tests/fixtures/word-implicit-final-2-sections-omitted.docx' : reference.rows[0].fixture;
  const data = Uint8Array.from(readFileSync(path)).buffer, source = await readDocx(data);
  const file = newFile('word', 'Invalid mark source', source.content); file.original = { name: 'source.docx', data };
  const editor = new Editor({ extensions: wordExtensions(), content: source.content.html });
  try {
    const key = crossPart ? [...source.paragraphs].find(([, p]) => p.ownerDocument !== source.documents.get('word/document.xml'))![0] : 'missing-source';
    editor.view.dispatch(editor.state.tr.setNodeAttribute(0, 'paragraphMarkSource', key));
    source.content.html = editor.getHTML(); const before = JSON.stringify(file), originalBytes = new Uint8Array(data).slice();
    await expect(exportRetainedDocument(file)).rejects.toThrow(crossPart ? /cross-part paragraph mark source/ : /unrecognized paragraph mark source/);
    expect(JSON.stringify(file)).toBe(before); expect(file.original.data).toBe(data);
    expect(new Uint8Array(file.original.data)).toEqual(originalBytes);
  } finally { editor.destroy(); }
});
