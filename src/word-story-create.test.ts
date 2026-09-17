import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { readDocx, wordXml, descendants, val, WORD_NS } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { sanitizeWordContent } from './formats';
import { newFile } from './model';
import { contentFingerprint } from './office-preservation';
import { resolveWordSections } from './word-section-layout';
import { createWordStory, changedWordStoryCreation } from './word-story-create';
import { changeWordStoryLink, wordStoryIsLinked } from './word-story-links';
import { WordStoryState, initializeWordStories, changeWordStory } from './word-story-edit';
import { wordExtensions } from './word-extensions';
import reference from '../tests/fixtures/native-word-story-create.json';
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
for (const [name, sample] of Object.entries(reference.cases))
  it(`creates and retains native active/inactive stories for ${name}`, async () => {
    const source = reference.sources[sample.mode as keyof typeof reference.sources];
    const data = Uint8Array.from(readFileSync('tests/fixtures/' + source.file)).buffer;
    expect(createHash('sha256').update(new Uint8Array(data)).digest('hex')).toBe(source.sha256);
    const content = sanitizeWordContent((await readDocx(data)).content);
    const file = newFile('word', 'Creation', content);
    file.original = {
      name: source.file,
      data,
      contentFingerprint: await contentFingerprint(content),
    };
    const originalZip = await JSZip.loadAsync(data);
    const originalStyles = await originalZip.file('word/styles.xml')!.async('string');
    const editor = new Editor({
      extensions: [...wordExtensions(), WordStoryState],
      content: content.html,
      parseOptions: { preserveWhitespace: true },
    });
    try {
      initializeWordStories(editor, content);
      const initial = editor.state.doc.toJSON();
      const sync = () => {
        content.stories = editor.state.doc.attrs.wordStories;
      };
      const sectionId = resolveWordSections(content)[sample.section - 1].id;
      const kind = sample.kind === 'Headers' ? 'header' : 'footer';
      const slot = ({ 1: 'default', 2: 'first', 3: 'even' } as const)[sample.slot as 1 | 2 | 3];
      const template = content.stories!.emptyTemplates![kind]!;
      expect(template).toContain('Calibri');
      expect(template).toContain('11pt');
      createWordStory(editor, content, content.stories, sectionId, kind, slot, template);
      expect(editor.state.doc.toJSON()).toEqual(initial);
      expect(editor.can().undo()).toBe(false);
      let actions = 0;
      if (sample.section === 2 && !('linked' in sample && sample.linked)) {
        changeWordStoryLink(editor, content, content.stories, sectionId, kind, slot, false);
        sync();
        actions++;
      }
      const inspect = async (stage: 'blank' | 'typed') => {
        const result = await bytes(await exportRetainedDocument(file));
        const reopened = (await readDocx(result)).content;
        expect(new DOMParser().parseFromString(reopened.html, 'text/html').body.textContent).toBe(
          sample[stage].body.replace(/[\r\n\v\f]/g, ''),
        );
        for (const observed of [content, reopened]) {
          for (const [index, section] of resolveWordSections(observed).entries())
            for (const [kind, key] of [
              ['header', 'headers'],
              ['footer', 'footers'],
            ] as const)
              for (const slot of ['default', 'first', 'even'] as const) {
                const expected = sample.references[stage][index][`${kind}:${slot}`];
                const ref = section[key][slot],
                  part =
                    ref &&
                    observed.stories!.parts.find((p) =>
                      p.relationshipIds.includes(ref.relationshipId),
                    );
                expect(
                  part
                    ? new DOMParser().parseFromString(part.html, 'text/html').body.textContent
                    : '',
                ).toBe(expected.text);
                expect(wordStoryIsLinked(observed, section.id, kind, slot)).toBe(expected.linked);
              }
        }
        const zip = await JSZip.loadAsync(result);
        for (const part of Object.values(originalZip.files))
          if (
            !part.dir &&
            ![
              'word/document.xml',
              'word/_rels/document.xml.rels',
              '[Content_Types].xml',
              'word/settings.xml',
              'word/styles.xml',
            ].includes(part.name)
          )
            expect(await zip.file(part.name)!.async('uint8array')).toEqual(
              await part.async('uint8array'),
            );
        if (sample.mode !== 'missing-styles')
          expect(await zip.file('word/styles.xml')!.async('string')).toBe(originalStyles);
      };
      await inspect('blank');
      const selected =
        resolveWordSections(content)[sample.section - 1][kind === 'header' ? 'headers' : 'footers'][
          slot
        ];
      const part =
        selected &&
        content.stories!.parts.find((p) => p.relationshipIds.includes(selected.relationshipId));
      const draft = new Editor({
        extensions: wordExtensions({ retainedEditRuns: true }),
        content: part?.html || template,
        parseOptions: { preserveWhitespace: true },
      });
      const before = editor.state.doc.toJSON();
      try {
        draft.commands.insertContent(sample.text);
        if (part) changeWordStory(editor, part.path, part.html, draft.getHTML());
        else
          createWordStory(editor, content, content.stories, sectionId, kind, slot, draft.getHTML());
      } finally {
        draft.destroy();
      }
      sync();
      actions++;
      const typed = editor.state.doc.toJSON();
      editor.commands.undo();
      expect(editor.state.doc.toJSON()).toEqual(before);
      editor.commands.redo();
      expect(editor.state.doc.toJSON()).toEqual(typed);
      sync();
      await inspect('typed');
      for (let i = 0; i < actions; i++) editor.commands.undo();
      sync();
      expect(editor.state.doc.toJSON()).toEqual(initial);
      expect(new Uint8Array(await bytes(await exportRetainedDocument(file)))).toEqual(
        new Uint8Array(data),
      );
      expect(file.original!.data).toEqual(data);
    } finally {
      editor.destroy();
    }
  });

it('keeps existing custom style identities while creating missing built-ins', async () => {
  const source = reference.sources['missing-styles'];
  const zip = await JSZip.loadAsync(readFileSync('tests/fixtures/' + source.file));
  const styles = wordXml(await zip.file('word/styles.xml')!.async('string'));
  for (const id of ['NofficeHeader', 'NofficeHeaderChar', 'NofficeFooter', 'NofficeFooterChar']) {
    const style = styles.createElementNS(WORD_NS, 'w:style');
    style.setAttributeNS(WORD_NS, 'w:type', 'paragraph');
    style.setAttributeNS(WORD_NS, 'w:styleId', id);
    style.setAttributeNS(WORD_NS, 'w:customStyle', '1');
    const name = styles.createElementNS(WORD_NS, 'w:name');
    name.setAttributeNS(WORD_NS, 'w:val', 'Existing ' + id);
    style.append(name);
    styles.documentElement.append(style);
  }
  const originalStyles = descendants(styles, 'style')
    .filter((s) => val(s, 'styleId').startsWith('Noffice'))
    .map((s) => new XMLSerializer().serializeToString(s));
  zip.file('word/styles.xml', new XMLSerializer().serializeToString(styles));
  const data = await zip.generateAsync({ type: 'arraybuffer' }),
    content = sanitizeWordContent((await readDocx(data)).content);
  const section = resolveWordSections(content)[0];
  content.stories = changedWordStoryCreation(
    content,
    section.id,
    'header',
    'default',
    content.stories!.emptyTemplates!.header!.replace('</p>', 'Created header</p>'),
  );
  const file = newFile('word', 'Collision', content);
  file.original = { name: 'collision.docx', data };
  const result = await JSZip.loadAsync(await bytes(await exportRetainedDocument(file)));
  const current = wordXml(await result.file('word/styles.xml')!.async('string'));
  expect(
    descendants(current, 'style')
      .filter((s) =>
        ['NofficeHeader', 'NofficeHeaderChar', 'NofficeFooter', 'NofficeFooterChar'].includes(
          val(s, 'styleId'),
        ),
      )
      .map((s) => new XMLSerializer().serializeToString(s)),
  ).toEqual(originalStyles);
  expect(
    descendants(current, 'style').filter((s) =>
      ['NofficeHeader1', 'NofficeHeaderChar1', 'NofficeFooter1', 'NofficeFooterChar1'].includes(
        val(s, 'styleId'),
      ),
    ),
  ).toHaveLength(4);
});

it('rejects a stale creation draft and forged paragraph identities atomically', async () => {
  const data = Uint8Array.from(
    readFileSync('tests/fixtures/' + reference.sources['existing-styles'].file),
  ).buffer;
  const content = sanitizeWordContent((await readDocx(data)).content),
    section = resolveWordSections(content)[0];
  const editor = new Editor({
    extensions: [...wordExtensions(), WordStoryState],
    content: content.html,
  });
  try {
    initializeWordStories(editor, content);
    const initial = editor.state.doc.toJSON();
    const html = content.stories!.emptyTemplates!.header!.replace('</p>', 'Created header</p>');
    const stale = { ...content.stories!, evenAndOddHeaders: !content.stories!.evenAndOddHeaders };
    expect(() =>
      createWordStory(editor, content, stale, section.id, 'header', 'default', html),
    ).toThrow('changed');
    expect(() =>
      createWordStory(
        editor,
        content,
        content.stories,
        section.id,
        'header',
        'default',
        html.replace('empty:0', '0:0'),
      ),
    ).toThrow('provenance');
    expect(editor.state.doc.toJSON()).toEqual(initial);
    expect(editor.can().undo()).toBe(false);
    const clean = changedWordStoryCreation(
      content,
      section.id,
      'header',
      'default',
      html.replace('</p>', '<script>bad()</script><img src="https://invalid.example/tracker"></p>'),
    );
    expect(JSON.stringify(clean)).not.toMatch(/<script\b|javascript:|invalid\.example|bad\(\)/i);
  } finally {
    editor.destroy();
  }
});
