import { it, expect, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { readDocx, wordXml, descendants, val, wElement } from './docx-import';
import { emptyStoryDocument, docxEmptyStoryTemplates } from './docx-empty-stories';
import { changedWordStoryCreation } from './word-story-create';
import { wordExtensions } from './word-extensions';
import { Editor } from '@tiptap/core';
import { resolveWordSections } from './word-section-layout';
import { exportRetainedDocument } from './docx-preserve';
import { newFile } from './model';
import { contentFingerprint } from './office-preservation';
import { hydrateWordStructure, needsWordStructure } from './word-structure';
import reference from '../tests/fixtures/native-word-story-base-styles.json';

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
const sourceBytes = (mode: keyof typeof reference.sources) =>
  Uint8Array.from(readFileSync('tests/fixtures/' + reference.sources[mode].file)).buffer;
const blobBytes = (blob: Blob) =>
  new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });

for (const sample of reference.cases)
  it(`restores missing style roots and exports native ${sample.name} without changing unrelated parts`, async () => {
    const data = sourceBytes(sample.mode as keyof typeof reference.sources);
    const { content } = await readDocx(data);
    const before = await JSZip.loadAsync(data);
    const file = newFile('word', 'Base styles', content);
    file.original = {
      name: 'source.docx',
      data,
      contentFingerprint: await contentFingerprint(content),
    };
    const kind = sample.kind === 'Headers' ? 'header' : 'footer';
    const editor = new Editor({
      extensions: wordExtensions(),
      content: content.stories!.emptyTemplates![kind]!,
    });
    try {
      editor.commands.insertContent(sample.text);
      content.stories = changedWordStoryCreation(
        content,
        resolveWordSections(content)[0].id,
        kind,
        'first',
        editor.getHTML(),
      );
    } finally {
      editor.destroy();
    }
    const result = await blobBytes(await exportRetainedDocument(file));
    const zip = await JSZip.loadAsync(result);
    const styles = wordXml(await zip.file('word/styles.xml')!.async('string'));
    for (const type of ['paragraph', 'character'])
      expect(
        descendants(styles, 'style').filter(
          (s) => val(s, 'type') === type && ['1', 'true', 'on'].includes(val(s, 'default')),
        ),
      ).toHaveLength(1);
    const restored = (await readDocx(result)).content;
    const section = resolveWordSections(restored)[0];
    const ref = (kind === 'header' ? section.headers : section.footers).first!;
    const part = restored.stories!.parts.find((p) =>
      p.relationshipIds.includes(ref.relationshipId),
    )!;
    const html = new DOMParser().parseFromString(part.html, 'text/html');
    expect(html.body.textContent).toBe(sample.text);
    const native = sample.typed.sections[0].stories.find(
      (s) => s.kind === sample.kind && s.slot === 2,
    )!.characters[0];
    expect(html.querySelector('span')!.style.fontSize).toBe(`${native.size}pt`);
    expect(html.querySelector('span')!.style.fontFamily.replace(/^["']|["']$/g, '')).toBe(
      native.font,
    );
    const settings = async (archive: JSZip) => {
      const xml = wordXml(await archive.file('word/settings.xml')!.async('string'));
      for (const rsids of descendants(xml, 'rsids')) rsids.remove();
      return new XMLSerializer().serializeToString(xml);
    };
    // Creation adds native editing-session IDs; other settings must be preserved.
    expect(await settings(zip)).toBe(await settings(before));
    for (const entry of Object.values(before.files)) {
      if (
        entry.dir ||
        [
          'word/styles.xml',
          'word/settings.xml',
          'word/document.xml',
          'word/_rels/document.xml.rels',
          '[Content_Types].xml',
        ].includes(entry.name)
      )
        continue;
      expect(await zip.file(entry.name)!.async('uint8array')).toEqual(
        await entry.async('uint8array'),
      );
    }
  });

it('migrates unavailable legacy templates without changing edited text, original identity or revision', async () => {
  const data = sourceBytes('missing-both');
  const { content } = await readDocx(data);
  content.stories!.templateVersion = 2;
  content.stories!.emptyTemplates = { header: null, footer: null };
  content.html = content.html.replace('B1-01', 'Saved edit');
  const file = newFile('word', 'Legacy styles', content);
  file.original = { name: 'source.docx', data, contentFingerprint: 'edited-original' };
  const before = structuredClone(file);
  expect(needsWordStructure(file)).toBe(true);
  const migrated = await hydrateWordStructure(file);
  if (migrated.content.kind !== 'word') throw Error('Expected Word');
  expect(migrated.content.html).toBe(content.html);
  expect(migrated.content.stories?.emptyTemplates?.header).toContain('11pt');
  expect(migrated.content.stories?.templateVersion).toBe(4);
  expect(migrated.original).toBe(file.original);
  expect(migrated.revision).toBe(file.revision);
  expect(await hydrateWordStructure(migrated)).toBe(migrated);
  expect(file).toEqual(before);
});

it('imports empty-story templates from a missing styles part without changing retained bytes', async () => {
  const data = sourceBytes('missing-styles-part');
  const zip = await JSZip.loadAsync(data);
  const templates = await docxEmptyStoryTemplates(zip);
  expect(templates.header).toContain('11pt');
  expect((await readDocx(data)).content.stories!.emptyTemplates).toEqual(templates);
  expect(zip.file('word/styles.xml')).toBeNull();
});

it('avoids occupied base-style IDs and rejects ambiguous defaults without mutating the package', async () => {
  const zip = await JSZip.loadAsync(sourceBytes('missing-both'));
  const styles = wordXml(await zip.file('word/styles.xml')!.async('string'));
  for (const id of ['NofficeNormal', 'NofficeDefaultParagraphFont']) {
    const style = wElement(styles, 'style', { type: 'character', styleId: id });
    style.append(wElement(styles, 'name', { val: 'Existing ' + id }));
    styles.documentElement.append(style);
  }
  zip.file('word/styles.xml', new XMLSerializer().serializeToString(styles));
  await emptyStoryDocument(zip, 'header');
  const result = wordXml(await zip.file('word/styles.xml')!.async('string'));
  expect(
    descendants(result, 'style')
      .filter((s) => val(s, 'default') === '1')
      .map((s) => val(s, 'styleId')),
  ).toEqual(expect.arrayContaining(['NofficeNormal1', 'NofficeDefaultParagraphFont1']));
  const duplicate = wElement(styles, 'style', {
    type: 'paragraph',
    default: '1',
    styleId: 'BaseA',
  });
  const duplicate2 = wElement(styles, 'style', {
    type: 'paragraph',
    default: '1',
    styleId: 'BaseB',
  });
  styles.documentElement.append(duplicate, duplicate2);
  const invalid = new XMLSerializer().serializeToString(styles);
  zip.file('word/styles.xml', invalid);
  await expect(emptyStoryDocument(zip, 'header')).rejects.toThrow('ambiguous default');
  expect(await zip.file('word/styles.xml')!.async('string')).toBe(invalid);
});

it('keeps unchanged legacy originals byte-identical after restoring their templates', async () => {
  const data = sourceBytes('missing-both');
  const { content } = await readDocx(data);
  content.stories!.templateVersion = 2;
  content.stories!.emptyTemplates = { header: null, footer: null };
  const file = newFile('word', 'Legacy unchanged', content);
  file.original = {
    name: 'source.docx',
    data,
    contentFingerprint: await contentFingerprint(content),
  };
  const migrated = await hydrateWordStructure(file);
  expect(new Uint8Array(await blobBytes(await exportRetainedDocument(migrated)))).toEqual(
    new Uint8Array(data),
  );
  expect(migrated.revision).toBe(file.revision);
});
