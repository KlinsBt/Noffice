import { expect, it, vi } from 'vitest';
import { Editor } from '@tiptap/core';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { readDocx, wordXml, descendants, val } from './docx-import';
import { wordExtensions } from './word-extensions';
import { resolveWordSections } from './word-section-layout';
import { newFile } from './model';
import { exportOffice } from './formats';
import native from '../tests/fixtures/native-word-script-scopes.json';
import history from '../tests/fixtures/native-word-script-history.json';
import replacement from '../tests/fixtures/native-word-paragraph-replacement.json';

vi.mock('mammoth', async original => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});

async function open(file: string, digest: string, kind: string, expectedText: string) {
  const bytes = readFileSync('tests/fixtures/word-script-scopes/' + file);
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(digest);
  const { content } = await readDocx(new Uint8Array(bytes).buffer);
  const section = resolveWordSections(content)[0];
  const relationship = section[kind === 'Headers' ? 'headers' : 'footers'].default;
  const html = kind === 'Body' ? content.html : content.stories!.parts.find(p =>
    p.relationshipIds.includes(relationship!.relationshipId))!.html;
  expect(new DOMParser().parseFromString(html, 'text/html').querySelector('p')?.textContent).toBe(expectedText);
  return new Editor({ extensions: wordExtensions(), content: html });
}

for (const row of history.rows) it(`matches native end-caret format Undo/Redo typing state: ${row.name}`, async () => {
  const reference = native.rows.find(r => r.name === row.name)!, source = native.sources.find(s => s.name === reference.source)!;
  const text = source.characters.map(c => c.text).join('').replace(/\r$/, '');
  const editor = await open(source.file, source.docxHash, reference.kind, text);
  try {
    editor.commands.setTextSelection(text.length + 1);
    editor.commands.toggleWordScript(reference.script === 'Superscript' ? 'superscript' : 'subscript');
    editor.commands.undo(); editor.commands.redo();
    expect(snapshot(editor)).toEqual(reference.formatted);
    editor.view.dispatch(editor.state.tr.insertText('X'));
    expect(snapshot(editor)).toEqual(row.typed);
  } finally { editor.destroy(); }
});

for (const row of replacement.rows) it(`matches baseline native paragraph replacement and history: ${row.name}`, async () => {
  const source = native.sources.find(s => s.file === row.source)!;
  const text = source.characters.map(c => c.text).join('').replace(/\r$/, '');
  const editor = await open(source.file, source.docxHash, row.kind, text);
  const story = () => editor.state.doc.content.content.map(p => p.textBetween(0, p.content.size, '', '\t') + '\r').join('');
  try {
    expect(story()).toBe(row.before);
    editor.commands.setTextSelection({ from: row.selection.start + 1, to: editor.state.doc.firstChild!.nodeSize + 1 });
    const before = editor.getJSON(), { from, to } = editor.state.selection;
    const handled = editor.view.someProp('handleTextInput', fn => fn(editor.view, from, to, 'X', () => editor.state.tr.insertText('X')));
    expect(handled).toBe(true); expect(story()).toBe(row.typed);
    editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
    editor.commands.redo(); expect(story()).toBe(row.redo);
  } finally { editor.destroy(); }
});
function snapshot(editor: Editor) {
  const p = editor.getJSON().content![0];
  const read = (text: string, marks: typeof p.marks = [], paragraph = false) => {
    const attrs = marks?.find(m => m.type === 'textStyle')?.attrs;
    return { text, family: String(attrs?.fontFamily || p.attrs!.paragraphFontFamily).replace(/^["']|["']$/g, ''),
      size: parseFloat(String(attrs?.fontSize || p.attrs!.paragraphFontSize)),
      superscript: (paragraph ? p.attrs!.paragraphScript === 'superscript' : marks?.some(m => m.type === 'superscript')) ? -1 : 0,
      subscript: (paragraph ? p.attrs!.paragraphScript === 'subscript' : marks?.some(m => m.type === 'subscript')) ? -1 : 0 };
  };
  return [...(p.content || []).flatMap(n => [...(n.type === 'wordTab' ? '\t' : 'text' in n ? n.text : '')].map(c => read(c, n.marks))), read('\r', [], true)];
}

for (const row of native.rows) {
  const source = native.sources.find(s => s.name === row.source)!;
  const text = source.characters.map(c => c.text).join('').replace(/\r$/, '');
  it(`imports native paragraph script and inline formatting: ${row.name}`, async () => {
    const editor = await open(row.formattedFile, row.stages[0].docxHash, row.kind, text);
    try { expect(snapshot(editor)).toEqual(row.formatted); } finally { editor.destroy(); }
  });
  it(`applies native script scope and atomic history: ${row.name}`, async () => {
    const editor = await open(source.file, source.docxHash, row.kind, text);
    try {
      expect(snapshot(editor)).toEqual(source.characters);
      const end = row.selection.end > text.length ? editor.state.doc.firstChild!.nodeSize + 1 : row.selection.end + 1;
      editor.commands.setTextSelection({ from: row.selection.start + 1, to: end });
      const before = editor.getJSON();
      editor.commands.toggleWordScript(row.script === 'Superscript' ? 'superscript' : 'subscript');
      expect(snapshot(editor)).toEqual(row.formatted);
      const formatted = editor.getJSON();
      editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
      editor.commands.redo(); expect(editor.getJSON()).toEqual(formatted);
    } finally { editor.destroy(); }
  });
  it(`types the native selection scope without consuming its paragraph mark: ${row.name}`, async () => {
    const editor = await open(source.file, source.docxHash, row.kind, text);
    try {
      const end = row.selection.end > text.length ? editor.state.doc.firstChild!.nodeSize + 1 : row.selection.end + 1;
      editor.commands.setTextSelection({ from: row.selection.start + 1, to: end });
      editor.commands.toggleWordScript(row.script === 'Superscript' ? 'superscript' : 'subscript');
      expect(snapshot(editor)).toEqual(row.formatted);
      const before = editor.getJSON(), { from, to } = editor.state.selection;
      const handled = editor.view.someProp('handleTextInput', fn => fn(editor.view, from, to, 'X', () => editor.state.tr.insertText('X')));
      if (!handled) editor.view.dispatch(editor.state.tr.insertText('X'));
      expect(snapshot(editor)).toEqual(row.typed);
      expect(editor.getJSON().content!.slice(1)).toEqual(before.content!.slice(1));
      const typed = editor.getJSON();
      editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
      editor.commands.redo(); expect(editor.getJSON()).toEqual(typed);
    } finally { editor.destroy(); }
  });
  it(`retains native scripts and unrelated parts in an edited DOCX: ${row.name}`, async () => {
    const bytes = readFileSync('tests/fixtures/word-script-scopes/' + source.file);
    const data = new Uint8Array(bytes).buffer;
    const { content } = await readDocx(data);
    const file = newFile('word', row.name, content);
    file.original = { name: 'source.docx', data };
    const editor = await open(source.file, source.docxHash, row.kind, text);
    const relationship = resolveWordSections(content)[0][row.kind === 'Headers' ? 'headers' : 'footers'].default;
    const part = row.kind === 'Body' ? null : content.stories!.parts.find(p => p.relationshipIds.includes(relationship!.relationshipId))!;
    try {
      const end = row.selection.end > text.length ? editor.state.doc.firstChild!.nodeSize + 1 : row.selection.end + 1;
      editor.commands.setTextSelection({ from: row.selection.start + 1, to: end });
      editor.commands.toggleWordScript(row.script === 'Superscript' ? 'superscript' : 'subscript');
      if (part) part.html = editor.getHTML(); else content.html = editor.getHTML();
      const output = await JSZip.loadAsync(await exportOffice(file));
      const original = await JSZip.loadAsync(data);
      for (const path of Object.keys(original.files)) {
        if (original.files[path].dir || [part?.path || 'word/document.xml', 'word/settings.xml'].includes(path)) continue;
        expect(await output.file(path)!.async('uint8array'), path).toEqual(await original.file(path)!.async('uint8array'));
      }
      const restored = (await readDocx(await output.generateAsync({ type: 'arraybuffer' }))).content;
      editor.commands.setContent(part ? restored.stories!.parts.find(p => p.path === part.path)!.html : restored.html);
      expect(snapshot(editor)).toEqual(row.formatted);
      if (row.kind === 'Body' && row.scope === 'text-selection') {
        editor.commands.setTextSelection({ from: 1, to: text.length + 1 });
        editor.view.dispatch(editor.state.tr.insertText('X'));
        content.html = editor.getHTML();
        const typedZip = await JSZip.loadAsync(await exportOffice(file));
        const typed = (await readDocx(await typedZip.generateAsync({ type: 'arraybuffer' }))).content;
        editor.commands.setContent(typed.html); expect(snapshot(editor)).toEqual(row.typed);
        const p = descendants(wordXml(await typedZip.file('word/document.xml')!.async('string')), 'p')[0];
        const start = [...p.children].findIndex(e => e.localName === 'bookmarkStart');
        expect(start).toBeGreaterThan([...p.children].findIndex(e => e.localName === 'r'));
        expect(val(p.children[start], 'name')).toBe('_GoBack');
      }
    } finally { editor.destroy(); }
  });
}
