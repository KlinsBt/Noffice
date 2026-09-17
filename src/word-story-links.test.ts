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
import { changedWordStoryLink, changeWordStoryLink } from './word-story-links';
import { WordStoryState, initializeWordStories } from './word-story-edit';
import { wordExtensions } from './word-extensions';
import reference from '../tests/fixtures/native-word-story-links.json';
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
async function setup(name: keyof typeof reference.cases = 'unlink-header-default') {
  const sample = reference.cases[name];
  const data = Uint8Array.from(readFileSync(`tests/fixtures/${sample.sourceFile}`)).buffer;
  expect(createHash('sha256').update(new Uint8Array(data)).digest('hex')).toBe(sample.sourceSha256);
  const read = await readDocx(data),
    content = sanitizeWordContent(read.content);
  const file = newFile('word', 'Links', content);
  file.original = {
    name: 'Links.docx',
    data,
    contentFingerprint: await contentFingerprint(content),
  };
  if (file.content.kind !== 'word') throw Error('Expected Word');
  return { file, content: file.content, data, sample };
}
for (const name of Object.keys(reference.cases) as (keyof typeof reference.cases)[])
  it(`retains unrelated parts and reimports all native story states after ${name}`, async () => {
    const { file, content, data, sample } = await setup(name);
    const action = sample.action as {
      kind: 'header' | 'footer';
      slot: 'default' | 'first' | 'even';
      linked: boolean;
    };
    const source = structuredClone(content);
    const id = resolveWordSections(content)[1].id;
    content.stories = changedWordStoryLink(content, id, action.kind, action.slot, action.linked);
    expect(content.docxStructure).toEqual(source.docxStructure);
    const result = await bytes(await exportRetainedDocument(file));
    const reopened = (await readDocx(result)).content;
    for (const [index, section] of resolveWordSections(reopened).entries()) {
      for (const native of sample.native.sections[index].stories) {
        const key = native.kind === 'Headers' ? 'headers' : 'footers';
        const slot = ({ 1: 'default', 2: 'first', 3: 'even' } as const)[native.slot as 1 | 2 | 3];
        const ref = section[key][slot];
        expect(!!ref?.inherited, `${index}:${key}:${slot} link`).toBe(native.linked);
        const part =
          ref &&
          reopened.stories!.parts.find((p) => p.relationshipIds.includes(ref.relationshipId));
        const text = part
          ? new DOMParser().parseFromString(part.html, 'text/html').body.textContent
          : '';
        expect(text, `${index}:${key}:${slot} text`).toBe(
          native.paragraphs.map((p) => p.text.replace(/\r/g, '')).join(''),
        );
      }
    }
    const before = await JSZip.loadAsync(data),
      after = await JSZip.loadAsync(result);
    for (const path of Object.keys(before.files))
      if (
        !before.files[path].dir &&
        ![
          'word/document.xml',
          'word/_rels/document.xml.rels',
          'word/settings.xml',
          '[Content_Types].xml',
        ].includes(path)
      )
        expect(await after.file(path)!.async('uint8array'), path).toEqual(
          await before.file(path)!.async('uint8array'),
        );
    expect(file.original!.data).toEqual(data);
  });

it('unlinks atomically, rejects stale drafts and restores source bytes through undo', async () => {
  const { content, file, data } = await setup();
  const id = resolveWordSections(content)[1].id;
  const editor = new Editor({
    extensions: [...wordExtensions(), WordStoryState],
    content: content.html,
  });
  try {
    initializeWordStories(editor, content);
    const original = structuredClone(content.stories);
    changeWordStoryLink(editor, content, content.stories, id, 'header', 'default', false);
    const detached = structuredClone(editor.state.doc.attrs.wordStories);
    expect(detached.parts.length).toBe(original!.parts.length + 1);
    expect(() =>
      changeWordStoryLink(editor, content, content.stories, id, 'header', 'default', true),
    ).toThrow(/changed/);
    editor.commands.undo();
    expect(editor.state.doc.attrs.wordStories).toEqual(original);
    content.stories = editor.state.doc.attrs.wordStories;
    expect(new Uint8Array(await bytes(await exportRetainedDocument(file)))).toEqual(
      new Uint8Array(data),
    );
    editor.commands.redo();
    expect(editor.state.doc.attrs.wordStories).toEqual(detached);
    content.stories = detached;
    content.stories = changedWordStoryLink(content, id, 'header', 'default', true);
    expect(content.stories).toEqual(original);
    expect(new Uint8Array(await bytes(await exportRetainedDocument(file)))).toEqual(
      new Uint8Array(data),
    );
  } finally {
    editor.destroy();
  }
});

it('keeps a linked successor attached to the copied part and removes unused copies when relinked', async () => {
  const { content } = await setup();
  const original = structuredClone(content.docxStructure!.sections[1]);
  content.docxStructure!.sections.push({ ...original, id: 'word/document.xml#section:2' });
  const id = resolveWordSections(content)[1].id;
  content.stories = changedWordStoryLink(content, id, 'header', 'default', false);
  let sections = resolveWordSections(content);
  expect(sections[1].headers.default!.relationshipId).not.toBe(
    sections[0].headers.default!.relationshipId,
  );
  expect(sections[2].headers.default).toEqual({ ...sections[1].headers.default, inherited: true });
  content.stories = changedWordStoryLink(content, id, 'header', 'default', true);
  sections = resolveWordSections(content);
  expect(sections[2].headers.default!.relationshipId).toBe(
    sections[0].headers.default!.relationshipId,
  );
  expect(content.stories.parts.some((p) => p.copiedFrom)).toBe(false);
});

it('keeps copied paragraph edits independent from the previous section across reloadable export', async () => {
  const { content, file } = await setup();
  content.stories = changedWordStoryLink(
    content,
    resolveWordSections(content)[1].id,
    'header',
    'default',
    false,
  );
  const part = content.stories.parts.find((p) => p.copiedFrom)!;
  const doc = new DOMParser().parseFromString(part.html, 'text/html');
  doc.querySelector('p')!.append(' independent');
  part.html = doc.body.innerHTML;
  const reopened = (await readDocx(await bytes(await exportRetainedDocument(file)))).content;
  const sections = resolveWordSections(reopened);
  const text = (index: number) => {
    const part = reopened.stories!.parts.find((p) =>
      p.relationshipIds.includes(sections[index].headers.default!.relationshipId),
    )!;
    return new DOMParser().parseFromString(part.html, 'text/html').body.textContent;
  };
  expect(text(0)).toBe('S1 header default');
  expect(text(1)).toBe('S1 header default independent');
});

it('rejects first-section links, invalid clone provenance and package collisions without changing originals', async () => {
  const { content, file, data } = await setup();
  expect(() =>
    changedWordStoryLink(content, resolveWordSections(content)[0].id, 'header', 'default', true),
  ).toThrow(/later section/);
  content.stories = changedWordStoryLink(
    content,
    resolveWordSections(content)[1].id,
    'header',
    'default',
    false,
  );
  const part = content.stories.parts.find((p) => p.copiedFrom)!;
  const source = part.copiedFrom;
  part.copiedFrom = '../missing.xml';
  await expect(exportRetainedDocument(file)).rejects.toThrow();
  part.copiedFrom = source;
  const zip = await JSZip.loadAsync(data);
  zip.file(part.path, '<opaque/>');
  file.original!.data = await zip.generateAsync({ type: 'arraybuffer' });
  await expect(exportRetainedDocument(file)).rejects.toThrow(/collides/);
  expect(new Uint8Array(data)).toEqual(new Uint8Array((await setup()).data));
});

it('copies the current header snapshot and keeps later source edits out of the copy', async () => {
  const { content, file } = await setup();
  const sourceId = resolveWordSections(content)[0].headers.default!.relationshipId;
  const append = (text: string) => {
    const part = content.stories!.parts.find((p) => p.relationshipIds.includes(sourceId))!;
    const doc = new DOMParser().parseFromString(part.html, 'text/html');
    doc.querySelector('p')!.append(text);
    part.html = doc.body.innerHTML;
  };
  append(' before');
  content.stories = changedWordStoryLink(
    content,
    resolveWordSections(content)[1].id,
    'header',
    'default',
    false,
  );
  append(' after');
  const read = (await readDocx(await bytes(await exportRetainedDocument(file)))).content;
  const texts = resolveWordSections(read).map((section) => {
    const part = read.stories!.parts.find((p) =>
      p.relationshipIds.includes(section.headers.default!.relationshipId),
    )!;
    return new DOMParser().parseFromString(part.html, 'text/html').body.textContent;
  });
  expect(texts).toEqual(['S1 header default before after', 'S1 header default before']);
});

it.each([false, true])(
  'preserves nested-source relationship targets or rejects package traversal (invalid=%s)',
  async (invalid) => {
    const initial = await setup(),
      zip = await JSZip.loadAsync(initial.data);
    const id = resolveWordSections(initial.content)[0].headers.default!.relationshipId;
    const part = initial.content.stories!.parts.find((p) => p.relationshipIds.includes(id))!;
    const header = await zip.file(part.path)!.async('string');
    zip.remove(part.path);
    zip.file('custom/header.xml', header, { createFolders: false });
    const rels = new DOMParser().parseFromString(
      await zip.file('word/_rels/document.xml.rels')!.async('string'),
      'application/xml',
    );
    [...rels.documentElement.children]
      .find((r) => r.getAttribute('Id') === id)!
      .setAttribute('Target', '../custom/header.xml');
    zip.file('word/_rels/document.xml.rels', new XMLSerializer().serializeToString(rels));
    const types = new DOMParser().parseFromString(
      await zip.file('[Content_Types].xml')!.async('string'),
      'application/xml',
    );
    [...types.documentElement.children]
      .find((r) => r.getAttribute('PartName') === `/${part.path}`)!
      .setAttribute('PartName', '/custom/header.xml');
    zip.file('[Content_Types].xml', new XMLSerializer().serializeToString(types));
    zip.file(
      'custom/_rels/header.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="external" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/preserved" TargetMode="External"/><Relationship Id="opaque" Type="urn:noffice:opaque" Target="${invalid ? '../../escape.xml' : 'assets/payload.xml'}"/></Relationships>`,
      { createFolders: false },
    );
    zip.file('custom/assets/payload.xml', '<opaque>preserved</opaque>', { createFolders: false });
    const data = await zip.generateAsync({ type: 'arraybuffer' });
    const content = (await readDocx(data)).content;
    const file = newFile('word', 'Nested source', content);
    file.original = {
      name: 'Nested.docx',
      data,
      contentFingerprint: await contentFingerprint(content),
    };
    content.stories = changedWordStoryLink(
      content,
      resolveWordSections(content)[1].id,
      'header',
      'default',
      false,
    );
    if (invalid) await expect(exportRetainedDocument(file)).rejects.toThrow(/leaves the package/);
    else {
      const output = await JSZip.loadAsync(await bytes(await exportRetainedDocument(file)));
      const copy = content.stories.parts.find((p) => p.copiedFrom)!;
      const relPath = copy.path.replace(/([^/]+)$/, '_rels/$1.rels');
      const refs = new DOMParser().parseFromString(
        await output.file(relPath)!.async('string'),
        'application/xml',
      );
      const values = [...refs.documentElement.children];
      expect(values.find((r) => r.getAttribute('Id') === 'external')!.getAttribute('Target')).toBe(
        'https://example.invalid/preserved',
      );
      expect(values.find((r) => r.getAttribute('Id') === 'opaque')!.getAttribute('Target')).toBe(
        '/custom/assets/payload.xml',
      );
      expect(await output.file('custom/header.xml')!.async('string')).toBe(header);
      expect(await output.file('custom/assets/payload.xml')!.async('string')).toBe(
        '<opaque>preserved</opaque>',
      );
    }
    expect(file.original.data).toEqual(data);
  },
);
