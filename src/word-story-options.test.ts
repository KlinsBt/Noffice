import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { readDocx } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { sanitizeWordContent } from './formats';
import { newFile } from './model';
import { contentFingerprint } from './office-preservation';
import { resolveWordSections } from './word-section-layout';
import {
  changedWordStoryOptions,
  changeWordStoryPageOptions,
  wordStoryDistancePoints,
} from './word-story-options';
import { WordStoryState, initializeWordStories } from './word-story-edit';
import { wordExtensions } from './word-extensions';
import reference from '../tests/fixtures/native-word-story-options.json';
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

const data = Uint8Array.from(readFileSync(`tests/fixtures/${reference.sourceFile}`)).buffer;
const blobBytes = (blob: Blob) =>
  new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });
async function setup() {
  expect(createHash('sha256').update(new Uint8Array(data)).digest('hex')).toBe(
    reference.sourceSha256,
  );
  const read = await readDocx(data);
  const file = newFile('word', 'Options', sanitizeWordContent(read.content));
  file.original = {
    name: 'Options.docx',
    data,
    contentFingerprint: await contentFingerprint(file.content),
  };
  if (file.content.kind !== 'word') throw Error('Expected Word');
  return { file, content: file.content };
}
for (const [name, sample] of Object.entries(reference.cases))
  it(`preserves unrelated package parts and reimports native ${name} page options`, async () => {
    const { file, content } = await setup();
    const action = sample.action as {
      first?: number[];
      even?: boolean;
      section?: number;
      distance?: string;
      value?: number;
    };
    for (let i = 0; i < 2; i++) {
      const section = resolveWordSections(content)[i];
      content.stories = changedWordStoryOptions(content, section.id, {
        differentFirstPage: action.first?.includes(i + 1) ? false : section.differentFirstPage,
        evenAndOddHeaders: action.even ? false : content.stories!.evenAndOddHeaders,
        headerDistance:
          action.section === i + 1 && action.distance === 'HeaderDistance'
            ? wordStoryDistancePoints(String(action.value))
            : section.margins.header!,
        footerDistance:
          action.section === i + 1 && action.distance === 'FooterDistance'
            ? wordStoryDistancePoints(String(action.value))
            : section.margins.footer!,
      });
    }
    const result = await blobBytes(await exportRetainedDocument(file));
    if (name === 'baseline') expect(new Uint8Array(result)).toEqual(new Uint8Array(data));
    const reimported = await readDocx(result);
    for (const [i, section] of resolveWordSections(reimported.content).entries()) {
      const native = sample.native.sections[i];
      expect(section.differentFirstPage).toBe(native.first);
      expect(reimported.content.stories!.evenAndOddHeaders).toBe(native.even);
      expect(section.margins.header).toBe(Math.round(native.header * 20));
      expect(section.margins.footer).toBe(Math.round(native.footer * 20));
    }
    const before = await JSZip.loadAsync(data),
      after = await JSZip.loadAsync(result);
    expect(Object.keys(after.files).sort()).toEqual(Object.keys(before.files).sort());
    for (const path of Object.keys(before.files))
      if (!before.files[path].dir && !['word/document.xml', 'word/settings.xml'].includes(path))
        expect(await after.file(path)!.async('uint8array'), path).toEqual(
          await before.file(path)!.async('uint8array'),
        );
  });
it('keeps page options atomic with history, rejects stale drafts and removes restored overrides', async () => {
  const { content } = await setup();
  const editor = new Editor({
    extensions: [...wordExtensions(), WordStoryState],
    content: content.html,
    parseOptions: { preserveWhitespace: true },
  });
  try {
    initializeWordStories(editor, content);
    const original = structuredClone(editor.state.doc.attrs.wordStories);
    const id = resolveWordSections(content)[1].id;
    const options = {
      differentFirstPage: false,
      evenAndOddHeaders: false,
      headerDistance: 600,
      footerDistance: 600,
    };
    changeWordStoryPageOptions(editor, content, content.stories, id, options);
    const edited = structuredClone(editor.state.doc.attrs.wordStories);
    expect(edited.sectionOptions).toHaveLength(1);
    expect(() => changeWordStoryPageOptions(editor, content, content.stories, id, options)).toThrow(
      /changed/,
    );
    editor.commands.undo();
    expect(editor.state.doc.attrs.wordStories).toEqual(original);
    editor.commands.redo();
    expect(editor.state.doc.attrs.wordStories).toEqual(edited);
    changeWordStoryPageOptions(editor, content, edited, id, {
      differentFirstPage: true,
      evenAndOddHeaders: true,
      headerDistance: 360,
      footerDistance: 360,
    });
    expect(editor.state.doc.attrs.wordStories).toEqual(original);
  } finally {
    editor.destroy();
  }
});
it('rejects invalid distances and foreign sections before changing content or exporting', async () => {
  const { file, content } = await setup();
  const before = structuredClone(content);
  const id = resolveWordSections(content)[0].id;
  for (const invalid of [-1, NaN, Infinity, 31681, 0.5])
    expect(() =>
      changedWordStoryOptions(content, id, {
        differentFirstPage: true,
        evenAndOddHeaders: true,
        headerDistance: invalid,
        footerDistance: 360,
      }),
    ).toThrow(/distances/);
  expect(content).toEqual(before);
  expect(() =>
    changedWordStoryOptions(content, 'word/document.xml#section:999', {
      differentFirstPage: true,
      evenAndOddHeaders: true,
      headerDistance: 360,
      footerDistance: 360,
    }),
  ).toThrow(/section/);
  content.stories!.sectionOptions = [
    { sectionId: 'word/document.xml#section:999', headerDistance: 360 },
  ];
  await expect(exportRetainedDocument(file)).rejects.toThrow(/unavailable source section/);
});
