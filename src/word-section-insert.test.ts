import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { readDocx } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { newFile, contentSchema } from './model';
import { contentFingerprint } from './office-preservation';
import { wordExtensions, wordJSON } from './word-extensions';
import {
  WordEditorSections,
  updateWordSectionSource,
  wordSectionMap,
} from './word-editor-sections';
import { WordStoryState, initializeWordStories } from './word-story-edit';
import { insertWordSectionBreak } from './word-section-insert';
import {
  sectionParagraphs,
  validateSectionState,
  type WordSectionState,
} from './word-section-breaks';
import { wordSectionStartSchema } from './word-section-identity';
import { resolveWordSections } from './word-section-layout';
import reference from '../tests/fixtures/word-section-insertion/reference.json';
import { exportOffice } from './formats';
import { changedWordStoryCreation } from './word-story-create';
import { changedWordStoryOptions } from './word-story-options';

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
const bytes = (blob: Blob) =>
  new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });
const extensions = () => [...wordExtensions(), WordEditorSections, WordStoryState];
for (const sample of reference.rows)
  it(`inserts ${sample.name} with native selection, history, sections and retained export`, async () => {
    const data = Uint8Array.from(readFileSync(sample.source)).buffer;
    expect(createHash('sha256').update(new Uint8Array(data)).digest('hex')).toBe(sample.sourceHash);
    const { content } = await readDocx(data),
      file = newFile('word', sample.name, content);
    file.original = {
      name: sample.name + '.docx',
      data,
      contentFingerprint: await contentFingerprint(content),
    };
    const originalModel = JSON.stringify(content);
    const editor = new Editor({
      extensions: extensions(),
      content: content.html,
      parseOptions: { preserveWhitespace: true },
    });
    try {
      updateWordSectionSource(editor, content.docxStructure);
      initializeWordStories(editor, content);
      const ps = sectionParagraphs(editor.state.doc);
      // Word positions count one paragraph/section separator; PM also counts the opening token.
      const position = (offset: number) => {
        let n = 0;
        for (const p of ps) {
          if (offset <= n + p.node.textContent.length) return p.from + 1 + offset - n;
          n += p.node.textContent.length + 1;
        }
        throw Error('Reference selection outside the source');
      };
      editor.commands.setTextSelection({
        from: position(sample.selection.from),
        to: position(sample.selection.to),
      });
      const before = editor.getJSON(),
        from = editor.state.selection.from;
      expect(
        insertWordSectionBreak(editor, content, wordSectionStartSchema.parse(sample.kind)),
      ).toBe(true);
      expect(editor.state.selection.from).toBe(from + 2);
      expect(editor.state.selection.empty).toBe(true);
      const after = editor.getJSON();
      const state = editor.state.doc.attrs.wordSectionState as WordSectionState;
      const boundaries = new Set(state.breaks.map((b) => b.paragraph));
      expect(
        sectionParagraphs(editor.state.doc)
          .map((p, i) => p.node.textContent + (boundaries.has(i) ? '\f' : '\r'))
          .join(''),
      ).toBe(sample.after);
      expect(wordSectionMap(editor.state)?.ranges).toHaveLength(sample.sections.length);
      editor.commands.undo();
      expect(editor.getJSON()).toEqual(before);
      editor.commands.redo();
      expect(editor.getJSON()).toEqual(after);
      const saved = {
        ...content,
        html: editor.getHTML(),
        sectionState: state,
        stories: editor.state.doc.attrs.wordStories || undefined,
      };
      expect(contentSchema.parse(saved)).toEqual(saved);
      const loaded = new Editor({
        extensions: extensions(),
        content: saved.html,
        parseOptions: { preserveWhitespace: true },
      });
      try {
        updateWordSectionSource(loaded, content.docxStructure, saved.sectionState);
        expect(wordSectionMap(loaded.state)).toEqual(wordSectionMap(editor.state));
      } finally {
        loaded.destroy();
      }
      expect(JSON.stringify(content)).toBe(originalModel);
      file.content = saved;
      const output = await bytes(await exportRetainedDocument(file));
      const imported = await readDocx(output);
      const layouts = resolveWordSections(imported.content);
      expect(layouts.map((s) => s.start)).toEqual(sample.types);
      layouts.forEach((s, i) => {
        const native = sample.sections[i];
        expect(s.width).toBeCloseTo(native.width * 20, 3);
        expect(s.height).toBeCloseTo(native.height * 20, 3);
        for (const side of ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const)
          expect(s.margins[side]).toBeCloseTo(native[side] * 20, 3);
        expect(s.differentFirstPage).toBe(!!native.first);
      });
      const a = await JSZip.loadAsync(data),
        b = await JSZip.loadAsync(output);
      for (const path of Object.keys(a.files).filter((p) =>
        /^word\/(header|footer).*\.xml$/.test(p),
      ))
        expect(await b.file(path)!.async('uint8array')).toEqual(
          await a.file(path)!.async('uint8array'),
        );
      expect(new Uint8Array(file.original.data)).toEqual(new Uint8Array(data));
    } finally {
      editor.destroy();
    }
  });

it('rejects nested selections and malformed authored provenance without changing content or history', () => {
  const content = {
    kind: 'word' as const,
    html: '<blockquote><p>Nested</p></blockquote><p>Body</p>',
    paper: 'a4' as const,
    margin: 'normal' as const,
  };
  const editor = new Editor({ extensions: extensions(), content: content.html });
  try {
    editor.commands.setTextSelection(3);
    const original = editor.getJSON();
    expect(insertWordSectionBreak(editor, content, 'nextPage')).toBe(false);
    expect(editor.getJSON()).toEqual(original);
    expect(editor.can().undo()).toBe(false);
    expect(() =>
      validateSectionState(
        {
          version: 2,
          finalSectionId: 'authored-body',
          breaks: [],
          inserted: [
            {
              id: 'authored-section:' + 'a'.repeat(32),
              sourceSectionId: 'word/document.xml#section:999',
            },
          ],
          starts: [],
        },
        undefined,
        editor.state.doc,
      ),
    ).toThrow(/provenance/);
  } finally {
    editor.destroy();
  }
});

it.each(['nextPage', 'continuous', 'evenPage', 'oddPage'] as const)(
  'authors fresh %s sections with stories, repeated insertion, deletion and export',
  async (start) => {
    const content = {
      ...newFile('word', 'Fresh').content,
      kind: 'word' as const,
      html: '<p>Alpha.</p><p>Beta</p>',
      paper: 'letter' as const,
      margin: 'normal' as const,
    };
    const stories = changedWordStoryCreation(
      content,
      'authored-body',
      'header',
      'default',
      '<p>Shared header</p>',
    );
    const editor = new Editor({ extensions: extensions(), content: content.html });
    try {
      initializeWordStories(editor, { ...content, stories });
      editor.commands.setTextSelection(4);
      const before = editor.getJSON();
      expect(insertWordSectionBreak(editor, { ...content, stories }, start)).toBe(true);
      const one = editor.getJSON();
      editor.commands.insertContent('X');
      editor.commands.undo();
      expect(editor.getJSON()).toEqual(one);
      expect(insertWordSectionBreak(editor, { ...content, stories }, 'oddPage')).toBe(true);
      const two = editor.getJSON();
      const saved = {
        ...content,
        html: editor.getHTML(),
        sectionState: editor.state.doc.attrs.wordSectionState as WordSectionState,
        stories: editor.state.doc.attrs.wordStories,
      };
      expect(resolveWordSections(saved).map((s) => s.start)).toEqual([
        'nextPage',
        start,
        'oddPage',
      ]);
      const file = newFile('word', 'Fresh sections', saved);
      const output = await bytes(await exportOffice(file));
      const imported = await readDocx(output),
        sections = resolveWordSections(imported.content);
      const importedDoc = editor.schema.nodeFromJSON(wordJSON(imported.content.html));
      expect(sectionParagraphs(importedDoc).map((p) => p.node.textContent)).toEqual(
        sectionParagraphs(editor.state.doc).map((p) => p.node.textContent),
      );
      expect(sections.map((s) => s.start)).toEqual(['nextPage', start, 'oddPage']);
      expect(sections.map((s) => s.headers.default?.inherited)).toEqual([false, true, true]);
      expect(imported.content.stories?.parts.some((p) => p.html.includes('Shared header'))).toBe(
        true,
      );
      editor.commands.undo();
      expect(editor.getJSON()).toEqual(one);
      editor.commands.undo();
      expect(editor.getJSON()).toEqual(before);
      editor.commands.redo();
      editor.commands.redo();
      expect(editor.getJSON()).toEqual(two);
      // Deleting the selected new separator retains the final identity and repairs story links.
      editor.commands.setTextSelection({ from: 4, to: 6 });
      expect(editor.commands.deleteWordBoundary()).toBe(true);
      const live = editor.state.doc.attrs.wordSectionState as WordSectionState;
      expect(live.breaks).toHaveLength(1);
      validateSectionState(live, undefined, editor.state.doc);
      editor.commands.undo();
      expect(editor.getJSON()).toEqual(two);
    } finally {
      editor.destroy();
    }
  },
);

it.each(['fresh', 'retained'])(
  'keeps %s section options independent through reset and actual export',
  async (mode) => {
    const source = Uint8Array.from(readFileSync(reference.rows[0].source)).buffer;
    const content =
      mode === 'retained'
        ? (await readDocx(source)).content
        : {
            ...newFile('word', 'Options').content,
            kind: 'word' as const,
            html: '<p>AlphaBeta</p>',
            paper: 'a4' as const,
            margin: 'normal' as const,
          };
    const file = newFile('word', 'Options', content);
    if (mode === 'retained')
      file.original = {
        name: 'Options.docx',
        data: source,
        contentFingerprint: await contentFingerprint(content),
      };
    const editor = new Editor({ extensions: extensions(), content: content.html });
    try {
      updateWordSectionSource(editor, content.docxStructure);
      initializeWordStories(editor, content);
      editor.commands.setTextSelection(4);
      insertWordSectionBreak(editor, content, 'nextPage');
      const saved = {
        ...content,
        html: editor.getHTML(),
        sectionState: editor.state.doc.attrs.wordSectionState as WordSectionState,
        stories: editor.state.doc.attrs.wordStories || undefined,
      };
      const [left, right] = resolveWordSections(saved);
      const defaults = {
        differentFirstPage: false,
        evenAndOddHeaders: false,
        headerDistance: left.margins.header!,
        footerDistance: left.margins.footer!,
      };
      const editedRight = {
        ...saved,
        stories: changedWordStoryOptions(saved, right.id, {
          ...defaults,
          differentFirstPage: true,
          headerDistance: 960,
        }),
      };
      const editedLeft = {
        ...editedRight,
        stories: changedWordStoryOptions(editedRight, left.id, {
          ...defaults,
          differentFirstPage: true,
          headerDistance: 1200,
        }),
      };
      const reset = {
        ...editedLeft,
        stories: changedWordStoryOptions(editedLeft, left.id, defaults),
      };
      expect(
        resolveWordSections(reset).map((s) => [s.differentFirstPage, s.margins.header]),
      ).toEqual([
        [false, defaults.headerDistance],
        [true, 960],
      ]);
      file.content = reset;
      const exported = await readDocx(await bytes(await exportOffice(file)));
      expect(
        resolveWordSections(exported.content).map((s) => [s.differentFirstPage, s.margins.header]),
      ).toEqual([
        [false, defaults.headerDistance],
        [true, 960],
      ]);
    } finally {
      editor.destroy();
    }
  },
);
