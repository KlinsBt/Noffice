import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { readDocx } from './docx-import';
import { wordExtensions, wordJSON } from './word-extensions';
import { exportOffice, importFile, sanitizeHTML } from './formats';
import { newFile, type WordContent } from './model';
import { hydrateWordParagraphScripts } from './word-script-migration';
import { hydrateWordStructure } from './word-structure';
import { clonedStoryHtml } from './word-story-provenance';
import { contentFingerprint } from './office-preservation';
import { wordSelectionScript, wordScriptValue } from './word-script';
import native from '../tests/fixtures/native-word-script-scopes.json';

vi.mock('mammoth', async original => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});

for (const script of ['superscript', 'subscript'] as const) {
  it(`types an imported empty ${script} mark, retaining explicit overrides, paste, history and reload`, () => {
    const initial = `<p data-word-paragraph-script="${script}"></p>`;
    const editor = new Editor({ extensions: wordExtensions(), content: initial });
    const marks = () => editor.state.doc.firstChild!.firstChild!.marks.map(m => m.type.name);
    try {
      expect(wordSelectionScript(editor.state)).toBe(script);
      editor.commands.insertContent('X'); expect(marks()).toContain(script);
      editor.commands.undo(); expect(editor.getText()).toBe('');
      editor.commands.redo(); expect(marks()).toContain(script);
      const html = sanitizeHTML(editor.getHTML()); editor.commands.setContent(html);
      expect(marks()).toContain(script);
      editor.commands.setContent(initial); editor.commands.setWordScript('baseline');
      editor.commands.insertContent('X'); expect(marks()).not.toContain(script);
      editor.commands.setContent(initial);
      editor.view.dispatch(editor.state.tr.insertText('P').setMeta('uiEvent', 'paste'));
      expect(marks()).not.toContain(script);
    } finally { editor.destroy(); }
  });
  it(`writes fresh ${script} marks separately from ordinary inline text in actual DOCX`, async () => {
    const file = newFile('word', 'Scripts');
    if (file.content.kind !== 'word') throw Error('Word fixture');
    file.content.html = `<p data-word-paragraph-script="${script}">Plain</p><p data-word-paragraph-script="${script}"></p><p data-word-paragraph-script="baseline">Baseline</p>`;
    const zip = await JSZip.loadAsync(await exportOffice(file));
    const result = (await readDocx(await zip.generateAsync({ type: 'arraybuffer' }))).content;
    const json = wordJSON(result.html);
    expect(json.content!.map(p => p.attrs!.paragraphScript)).toEqual([script, script, 'baseline']);
    expect(json.content![0].content![0].marks?.some(m => m.type === script)).toBeFalsy();
  });
}

it('rejects invalid commands atomically and does not create empty undo events', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p>Before</p>' });
  try {
    editor.commands.setTextSelection({ from: 1, to: 7 });
    for (const value of ['invalid', '', 1, null, {}, 'SuperScript']) {
      expect(wordScriptValue(value)).toBeNull();
      expect(editor.commands.setWordScript(value as 'baseline')).toBe(false);
      expect(editor.can().setWordScript(value as 'baseline')).toBe(false);
    }
    expect(editor.can().undo()).toBe(false);
    expect(editor.can().setWordScript('superscript')).toBe(true);
    expect(editor.can().undo()).toBe(false);
    editor.commands.setWordScript('superscript');
    const changed = editor.getJSON();
    editor.commands.setWordScript('superscript');
    expect(editor.getJSON()).toEqual(changed);
    editor.commands.undo(); expect(editor.getHTML()).toBe('<p>Before</p>');
    expect(editor.can().undo()).toBe(false);
  } finally { editor.destroy(); }
});

const content = (html: string): WordContent => ({ kind: 'word', html, paper: 'a4', margin: 'normal' });
it('restores legacy paragraph marks while preserving edited inline scripts and immutable input', () => {
  const source = content('<p data-source-paragraph="0:0" data-word-paragraph-script="superscript">Original</p>');
  const saved = content('<p data-source-paragraph="0:0"><sub>Edited</sub> plain</p>');
  const before = structuredClone(saved), result = hydrateWordParagraphScripts(saved, source);
  expect(saved).toEqual(before);
  expect(result.paragraphScriptsVersion).toBe(1);
  expect(result.html).toContain('<sub>Edited</sub> plain');
  expect(wordJSON(result.html).content![0].attrs!.paragraphScript).toBe('superscript');
  const ambiguous = content(saved.html + '<p>Split</p>');
  const original = structuredClone(ambiguous);
  expect(() => hydrateWordParagraphScripts(ambiguous, source)).toThrow(/backup/);
  expect(ambiguous).toEqual(original);
  const baseline = content(source.html.replace('superscript', 'baseline'));
  expect(wordJSON(hydrateWordParagraphScripts(ambiguous, baseline).html).content!.map(p => p.attrs!.paragraphScript))
    .toEqual(['baseline', 'baseline']);
});

it('restores cloned and created story marks through exact provenance', () => {
  const source = content('<p data-source-paragraph="0:0"></p>');
  const original = { path: 'word/footer.xml', kind: 'footer' as const, relationshipIds: ['rFooter'],
    html: '<p data-source-paragraph="3:0" data-word-paragraph-script="subscript">Native</p>' };
  const template = '<p data-source-paragraph="empty:0" data-word-paragraph-script="superscript"></p>';
  source.stories = { version: 1, evenAndOddHeaders: false, parts: [original], emptyTemplates: { header: null, footer: template } };
  const missing = (html: string) => html.replace(/ data-word-paragraph-script="[a-z]+"/g, '');
  const copy = 'word/footerCopy.xml', created = 'word/footerCreated.xml';
  const saved: WordContent = { ...source, stories: { ...source.stories, parts: [
    { ...original, path: copy, copiedFrom: original.path, html: missing(clonedStoryHtml(original, copy)).replace('Native', '<sup>Edited</sup>') },
    { path: created, kind: 'footer', relationshipIds: ['rNew'], created: true,
      html: missing(clonedStoryHtml({ ...original, html: template }, created)).replace('</p>', 'New</p>') },
  ] } };
  const result = hydrateWordParagraphScripts(saved, source);
  expect(result.stories!.parts.map(p => wordJSON(p.html).content![0].attrs!.paragraphScript)).toEqual(['subscript', 'superscript']);
  expect(result.stories!.parts[0].html).toContain('<sup>Edited</sup>');
  expect(result.stories!.parts[1].html).toContain('New</p>');
});

it('hydrates unchanged retained native marks without changing revision or original export bytes', async () => {
  const row = native.rows.find(r => r.kind === 'Headers' && r.script === 'Superscript' && r.scope === 'mark-only')
    || native.rows.find(r => r.name === 'Headers-Superscript-mark-only')!;
  const bytes = readFileSync('tests/fixtures/word-script-scopes/' + row.formattedFile);
  const data = new Uint8Array(bytes).buffer;
  const file = await importFile(Object.assign(new File([data], 'Native.docx'), { arrayBuffer: async () => data }));
  if (file.content.kind !== 'word') throw Error('Word fixture');
  file.content = (await readDocx(new Uint8Array(bytes).buffer, { legacyParagraphScripts: true })).content;
  file.original!.contentFingerprint = await contentFingerprint(file.content);
  const before = structuredClone(file), result = await hydrateWordStructure(file);
  expect(file).toEqual(before); expect(result.revision).toBe(file.revision);
  expect(result.content.kind === 'word' && result.content.paragraphScriptsVersion).toBe(1);
  const output = await exportOffice(result);
  const resultBytes = await new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader(); reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = reject; reader.readAsArrayBuffer(output);
  });
  expect(Buffer.from(resultBytes)).toEqual(bytes);
});
