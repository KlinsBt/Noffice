import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { readDocx } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { newFile, contentSchema, type WordContent } from './model';
import { contentFingerprint } from './office-preservation';
import { wordExtensions } from './word-extensions';
import {
  WordEditorSections,
  updateWordSectionSource,
  wordSectionMap,
} from './word-editor-sections';
import { WordStoryState, initializeWordStories } from './word-story-edit';
import { WordPageLayout, initializeWordPageLayout, changeWordPageLayout } from './word-page-layout';
import {
  changeWordSectionLayout,
  type WordSectionLayoutChange,
} from './word-section-scoped-layout';
import { resolveWordSections } from './word-section-layout';
import { insertWordSectionBreak } from './word-section-insert';
import { exportOffice } from './formats';
import reference from '../tests/fixtures/word-section-scoped-layout/reference.json';
import { wordSectionGeometrySchema } from './word-section-geometry';
import { validateWordSectionIdentities } from './word-section-breaks';

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

const extensions = () => [...wordExtensions(), WordEditorSections, WordStoryState, WordPageLayout];
const changes = (operation: string): WordSectionLayoutChange[] => {
  switch (operation) {
    case 'margin-normal':
      return [{ key: 'margin', value: 'normal' }];
    case 'margin-wide':
      return [{ key: 'margin', value: 'wide' }];
    case 'a4':
      return [{ key: 'paper', value: 'a4' }];
    case 'start-next':
      return [{ key: 'start', value: 'nextPage' }];
    case 'margin-narrow':
      return [{ key: 'margin', value: 'narrow' }];
    case 'portrait':
      return [{ key: 'orientation', value: 'portrait' }];
    case 'landscape':
      return [{ key: 'orientation', value: 'landscape' }];
    case 'letter':
      return [{ key: 'paper', value: 'letter' }];
    case 'start-odd':
      return [{ key: 'start', value: 'oddPage' }];
    case 'start-even':
      return [{ key: 'start', value: 'evenPage' }];
    case 'start-continuous':
      return [{ key: 'start', value: 'continuous' }];
    case 'landscape-letter':
      return [...changes('landscape'), ...changes('letter')];
    case 'letter-landscape':
      return [...changes('letter'), ...changes('landscape')];
    default:
      throw Error('Unknown native operation');
  }
};
const savedContent = (editor: Editor, content: WordContent): WordContent => ({
  ...content,
  ...editor.state.doc.attrs.wordPageLayout,
  html: editor.getHTML(),
  sectionState: editor.state.doc.attrs.wordSectionState || undefined,
  stories: editor.state.doc.attrs.wordStories || undefined,
});
function compare(
  content: WordContent,
  native: (typeof reference.rows)[number]['native']['edited'],
) {
  const sections = resolveWordSections(content);
  expect(sections).toHaveLength(native.sections.length);
  sections.forEach((section, i) => {
    const expected = native.sections[i];
    expect(section.width).toBe(Math.round(expected.width * 20));
    expect(section.height).toBe(Math.round(expected.height * 20));
    expect(section.orientation).toBe(expected.orientation ? 'landscape' : 'portrait');
    expect(section.start).toBe(
      ({ 0: 'continuous', 2: 'nextPage', 3: 'evenPage', 4: 'oddPage' } as Record<number, string>)[
        expected.start
      ],
    );
    for (const side of ['left', 'right', 'top', 'bottom', 'header', 'footer', 'gutter'] as const)
      expect(section.margins[side]).toBe(Math.round(expected[side] * 20));
    expect(section.differentFirstPage).toBe(!!expected.differentFirst);
  });
}
for (const sample of reference.rows)
  it(`matches native scoped geometry, history and retained files: ${sample.name}`, async () => {
    const data = Uint8Array.from(readFileSync(sample.fixture)).buffer;
    expect(createHash('sha256').update(new Uint8Array(data)).digest('hex')).toBe(sample.sourceHash);
    const { content } = await readDocx(data);
    const original = JSON.stringify(content);
    const file = newFile('word', sample.name, content);
    file.original = {
      name: sample.name + '.docx',
      data,
      contentFingerprint: await contentFingerprint(content),
    };
    const editor = new Editor({
      extensions: extensions(),
      content: content.html,
      parseOptions: { preserveWhitespace: true },
    });
    try {
      updateWordSectionSource(editor, content.docxStructure);
      initializeWordStories(editor, content);
      initializeWordPageLayout(editor, content);
      const range = wordSectionMap(editor.state)!.ranges[Math.max(1, sample.target) - 1];
      editor.commands.setTextSelection(range.from + 1);
      const before = editor.getJSON(),
        selection = editor.state.selection.toJSON();
      let changed = 0;
      for (const change of changes(sample.operation)) {
        if (sample.target === 0 && change.key !== 'start') {
          const previous = JSON.stringify(editor.getJSON());
          changeWordPageLayout(editor, change.key, change.value, content);
          changed += Number(previous !== JSON.stringify(editor.getJSON()));
        } else changed += Number(changeWordSectionLayout(editor, content, change));
      }
      expect(changed === 0).toBe(sample.native.unchanged);
      expect(editor.state.selection.toJSON()).toEqual(selection);
      const after = editor.getJSON(),
        saved = savedContent(editor, content);
      compare(saved, sample.native.edited);
      expect(contentSchema.parse(saved)).toEqual(saved);
      for (let i = 0; i < changed; i++) expect(editor.commands.undo()).toBe(true);
      expect(editor.getJSON()).toEqual(before);
      for (let i = 0; i < changed; i++) expect(editor.commands.redo()).toBe(true);
      expect(editor.getJSON()).toEqual(after);
      file.content = saved;
      const output = await bytes(await exportRetainedDocument(file));
      compare((await readDocx(output)).content, sample.native.edited);
      const a = await JSZip.loadAsync(data),
        b = await JSZip.loadAsync(output);
      for (const path of Object.keys(a.files).filter(
        (p) => !a.files[p].dir && p !== 'word/document.xml',
      ))
        expect(await b.file(path)!.async('uint8array'), path).toEqual(
          await a.file(path)!.async('uint8array'),
        );
      expect(JSON.stringify(content)).toBe(original);
      expect(new Uint8Array(file.original.data)).toEqual(new Uint8Array(data));
    } finally {
      editor.destroy();
    }
  });

it('preserves fresh scoped geometry through insertion, global controls, history and the fresh writer', async () => {
  const file = newFile('word');
  if (file.content.kind !== 'word') throw Error();
  const content = file.content;
  const editor = new Editor({ extensions: extensions(), content: content.html });
  try {
    initializeWordStories(editor, content);
    initializeWordPageLayout(editor, content);
    expect(
      changeWordSectionLayout(editor, content, { key: 'orientation', value: 'landscape' }),
    ).toBe(true);
    expect(insertWordSectionBreak(editor, savedContent(editor, content), 'nextPage')).toBe(true);
    let saved = savedContent(editor, content);
    expect(resolveWordSections(saved).map((s) => s.orientation)).toEqual([
      'landscape',
      'landscape',
    ]);
    expect(changeWordSectionLayout(editor, saved, { key: 'paper', value: 'letter' })).toBe(true);
    saved = savedContent(editor, content);
    const before = editor.getJSON();
    changeWordPageLayout(editor, 'orientation', 'portrait', saved);
    expect(
      resolveWordSections(savedContent(editor, content)).map((s) => [s.width, s.height]),
    ).toEqual([
      [11906, 16838],
      [12240, 15840],
    ]);
    editor.commands.undo();
    expect(editor.getJSON()).toEqual(before);
    editor.commands.redo();
    file.content = savedContent(editor, content);
    const output = await bytes(await exportOffice(file));
    expect(
      resolveWordSections((await readDocx(output)).content).map((s) => [
        s.width,
        s.height,
        s.orientation,
      ]),
    ).toEqual([
      [11906, 16838, 'portrait'],
      [12240, 15840, 'portrait'],
    ]);
    const settled = editor.getJSON();
    expect(() =>
      changeWordSectionLayout(
        editor,
        file.content as WordContent,
        { key: 'paper', value: 'bad' } as unknown as WordSectionLayoutChange,
      ),
    ).toThrow();
    expect(editor.getJSON()).toEqual(settled);
    editor.commands.setTextSelection({ from: 1, to: editor.state.doc.content.size - 1 });
    expect(() =>
      changeWordSectionLayout(editor, file.content as WordContent, {
        key: 'margin',
        value: 'narrow',
      }),
    ).toThrow(/one section/);
    expect(editor.getJSON()).toEqual(settled);
  } finally {
    editor.destroy();
  }
});

it('rejects invalid saved geometry and foreign or duplicate section overrides', () => {
  const geometry = {
    sectionId: 'authored-body',
    width: 12240,
    height: 15840,
    orientation: 'portrait',
    margins: { left: 720, right: 720, top: 720, bottom: 720 },
  } as const;
  for (const invalid of [
    { ...geometry, width: Infinity },
    { ...geometry, height: 0 },
    { ...geometry, width: 31681 },
    { ...geometry, margins: { ...geometry.margins, left: -1 } },
    { ...geometry, margins: { ...geometry.margins, left: 12240 } },
  ])
    expect(wordSectionGeometrySchema.safeParse(invalid).success).toBe(false);
  const state = {
    version: 2 as const,
    finalSectionId: 'authored-body',
    breaks: [],
    inserted: [],
    starts: [],
    layouts: [geometry],
  };
  expect(validateWordSectionIdentities(state, undefined)).toEqual(['authored-body']);
  expect(() =>
    validateWordSectionIdentities({ ...state, layouts: [geometry, geometry] }, undefined),
  ).toThrow(/geometry override/);
  expect(() =>
    validateWordSectionIdentities(
      { ...state, layouts: [{ ...geometry, sectionId: 'word/document.xml#section:0' }] },
      undefined,
    ),
  ).toThrow(/geometry override/);
});
