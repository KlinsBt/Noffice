import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { readDocx } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { sanitizeWordContent } from './formats';
import { newFile } from './model';
import { contentFingerprint } from './office-preservation';
import { resolveWordSections } from './word-section-layout';
import { changeWordStoryLink } from './word-story-links';
import { WordStoryState, initializeWordStories, changeWordStory } from './word-story-edit';
import { WordEditorSections, updateWordSectionSource } from './word-editor-sections';
import { wordExtensions } from './word-extensions';
import reference from '../tests/fixtures/native-word-story-successors.json';
import leading from '../tests/fixtures/native-word-story-leading-deletion.json';
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
for (const [name, sample] of Object.entries({ ...reference.cases, ...leading.cases }))
  it(`matches native inherited/copy contents and history for ${name}`, async () => {
    const data = Uint8Array.from(readFileSync(`tests/fixtures/${reference.sourceFile}`)).buffer;
    expect(createHash('sha256').update(new Uint8Array(data)).digest('hex')).toBe(
      reference.sourceSha256,
    );
    const content = sanitizeWordContent((await readDocx(data)).content);
    const file = newFile('word', 'Successors', content);
    file.original = {
      name: 'Successors.docx',
      data,
      contentFingerprint: await contentFingerprint(content),
    };
    const editor = new Editor({
      extensions: [
        ...wordExtensions(),
        WordStoryState,
        WordEditorSections.configure({ source: content.docxStructure }),
      ],
      content: content.html,
      parseOptions: { preserveWhitespace: true },
    });
    try {
      updateWordSectionSource(editor, content.docxStructure);
      initializeWordStories(editor, content);
      const sync = () => {
        content.html = editor.getHTML();
        content.stories = editor.state.doc.attrs.wordStories;
        content.sectionState = editor.state.doc.attrs.wordSectionState;
      };
      sync();
      const initial = editor.state.doc.toJSON();
      let actions = 0;
      const link = (
        section: number,
        kind: 'header' | 'footer',
        linked: boolean,
        slot: 'default' | 'first' | 'even' = 'default',
      ) => {
        const id = resolveWordSections(content)[section - 1].id;
        changeWordStoryLink(editor, content, content.stories, id, kind, slot, linked);
        sync();
        actions++;
      };
      const edit = (section: number, kind: 'header' | 'footer', text: string) => {
        const ref =
          resolveWordSections(content)[section - 1][kind === 'header' ? 'headers' : 'footers']
            .default!;
        const part = content.stories!.parts.find((p) =>
          p.relationshipIds.includes(ref.relationshipId),
        )!;
        const doc = new DOMParser().parseFromString(part.html, 'text/html');
        const walker = doc.createTreeWalker(doc.querySelector('p')!, NodeFilter.SHOW_TEXT);
        const leaf = walker.nextNode()!;
        leaf.textContent = text;
        expect(walker.nextNode()).toBeNull();
        changeWordStory(editor, part.path, part.html, doc.body.innerHTML);
        sync();
        actions++;
      };
      if (name in leading.cases) {
        if (name.includes('copy') || name === 'header-edit-first') {
          link(
            2,
            name.startsWith('footer') ? 'footer' : 'header',
            false,
            name.startsWith('first-header')
              ? 'first'
              : name.startsWith('even-header')
                ? 'even'
                : 'default',
          );
          if (name === 'header-edit-first') edit(2, 'header', 'Updated header');
        }
        const before = editor.state.doc.toJSON();
        const first = editor.state.doc.child(0).nodeSize;
        const middle = first + editor.state.doc.child(1).nodeSize;
        const boundary = name.includes('boundary');
        const to = name.includes('two') || name.includes('middle') ? middle : first;
        const from = boundary ? to - 1 : name.includes('partial') ? 7 : 0;
        editor.view.dispatch(closeHistory(editor.state.tr.delete(from, boundary ? to + 1 : to)));
        editor.view.dispatch(closeHistory(editor.state.tr));
        sync();
        actions++;
        const deleted = editor.state.doc.toJSON();
        editor.commands.undo();
        expect(editor.state.doc.toJSON()).toEqual(before);
        editor.commands.redo();
        expect(editor.state.doc.toJSON()).toEqual(deleted);
        sync();
      } else if (name !== 'baseline') {
        const kind = name === 'unlink-middle-footer-edit' ? 'footer' : 'header';
        link(2, kind, false);
        if (name !== 'unlink-middle-header') edit(2, kind, `Updated ${kind}`);
        if (name === 'unlink-last-from-copy') {
          link(3, 'header', false);
          edit(3, 'header', 'Last header');
        }
        if (name === 'relink-middle') link(2, 'header', true);
        if (name.startsWith('delete-')) {
          const before = editor.state.doc.toJSON();
          const first = editor.state.doc.child(0).nodeSize;
          const from = name === 'delete-first-after-copy' ? 0 : first;
          const to =
            name === 'delete-first-after-copy' ? first : first + editor.state.doc.child(1).nodeSize;
          editor.view.dispatch(closeHistory(editor.state.tr.delete(from, to)));
          editor.view.dispatch(closeHistory(editor.state.tr));
          sync();
          actions++;
          expect(resolveWordSections(content)).toHaveLength(2);
          const deleted = editor.state.doc.toJSON();
          editor.commands.undo();
          expect(editor.state.doc.toJSON()).toEqual(before);
          editor.commands.redo();
          expect(editor.state.doc.toJSON()).toEqual(deleted);
          sync();
        }
      }
      const output = await bytes(await exportRetainedDocument(file));
      const reopened = (await readDocx(output)).content;
      const sections = resolveWordSections(reopened);
      expect(sections).toHaveLength(sample.native.sections.length);
      for (const [index, section] of sections.entries())
        for (const native of sample.native.sections[index].stories) {
          const key = native.kind === 'Headers' ? 'headers' : 'footers';
          const slot = ({ 1: 'default', 2: 'first', 3: 'even' } as const)[native.slot as 1 | 2 | 3];
          const ref = section[key][slot];
          expect(!!ref?.inherited, `${index}:${key}:${slot}`).toBe(native.linked);
          const part =
            ref &&
            reopened.stories!.parts.find((p) => p.relationshipIds.includes(ref.relationshipId));
          expect(
            part ? new DOMParser().parseFromString(part.html, 'text/html').body.textContent : '',
          ).toBe(native.paragraphs.map((p) => p.text.replace(/\r/g, '')).join(''));
        }
      const bodyText = new DOMParser().parseFromString(reopened.html, 'text/html').body.textContent;
      expect(bodyText).toBe(sample.native.body.replace(/[\r\n\v\f]/g, ''));
      for (let i = 0; i < actions; i++) editor.commands.undo();
      expect(editor.state.doc.toJSON()).toEqual(initial);
      sync();
      expect(new Uint8Array(await bytes(await exportRetainedDocument(file)))).toEqual(
        new Uint8Array(data),
      );
      expect(file.original!.data).toEqual(data);
    } finally {
      editor.destroy();
    }
  });
