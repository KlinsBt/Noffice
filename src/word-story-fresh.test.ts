import { it, expect, vi } from 'vitest';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { readDocx, descendants, wordXml, child, val } from './docx-import';
import { newFile, contentSchema } from './model';
import { exportOffice } from './formats';
import { emptyWordParagraph } from './word-defaults';
import { authoredWordStoryTemplates } from './word-authored-stories';
import { changedWordStoryCreation, createWordStory } from './word-story-create';
import { wordExtensions } from './word-extensions';
import { WordStoryState, initializeWordStories, wordStoryChoices } from './word-story-edit';
import { changedWordStoryOptions, changeWordStoryPageOptions } from './word-story-options';
import { resolveWordSections } from './word-section-layout';
import reference from '../tests/fixtures/native-word-story-fresh.json';
import sourceReceipt from '../tests/fixtures/native-word-story-fresh-source.json';

it('ships only the bundled Inter embedding while retaining the native fixture content', async () => {
  const bytes = await readFile('tests/fixtures/word-story-fresh-source.docx');
  const hash = (value: Uint8Array) => createHash('sha256').update(value).digest('hex');
  expect(hash(bytes)).toBe(sourceReceipt.fixtureHash);
  expect(sourceReceipt.nativeSourceHash).toBe(reference.sourceHash);
  const zip = await JSZip.loadAsync(bytes);
  expect(Object.keys(zip.files).filter((name) => name.startsWith('word/fonts/') && !zip.files[name].dir))
    .toEqual(sourceReceipt.embeddedFontParts);
  const table = wordXml(await zip.file('word/fontTable.xml')!.async('string'));
  const embedded = descendants(table, 'font').filter((font) => child(font, 'embedRegular'));
  expect(embedded.map((font) => val(font, 'name'))).toEqual(['Inter']);
  for (const [name, expected] of Object.entries(sourceReceipt.preservedParts))
    expect(hash(await zip.file(name)!.async('uint8array')), name).toBe(expected);
});
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
const blobBytes = (blob: Blob) =>
  new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });
function fresh() {
  const file = newFile('word', 'Fresh stories');
  if (file.content.kind !== 'word') throw Error('Expected Word');
  file.content.html = emptyWordParagraph.replace(
    '</p>',
    Array.from({ length: 70 }, (_, i) => 'line' + String(i + 1).padStart(2, '0')).join('<br>') +
      '</p>',
  );
  return { file, content: file.content };
}
for (const sample of reference.rows)
  it(`exports fresh ${sample.name} with native defaults, six independent slots and preserved body text`, async () => {
    const { file, content } = fresh();
    const kind = sample.kind === 'Headers' ? 'header' : 'footer';
    const slot = ({ 1: 'default', 2: 'first', 3: 'even' } as const)[sample.slot as 1 | 2 | 3];
    expect(wordStoryChoices(content)).toHaveLength(6);
    content.stories = changedWordStoryOptions(content, 'authored-body', {
      differentFirstPage: slot === 'first',
      evenAndOddHeaders: slot === 'even',
      headerDistance: 708,
      footerDistance: 708,
    });
    const editor = new Editor({
      extensions: wordExtensions(),
      content: authoredWordStoryTemplates(content)![kind],
    });
    try {
      editor.commands.insertContent(sample.text);
      content.stories = changedWordStoryCreation(
        content,
        'authored-body',
        kind,
        slot,
        editor.getHTML(),
      );
    } finally {
      editor.destroy();
    }
    expect(content.stories.parts).toHaveLength(6);
    expect(
      new Set(content.stories.references!.filter((r) => !r.linked).map((r) => r.relationshipId))
        .size,
    ).toBe(6);
    file.content = contentSchema.parse(JSON.parse(JSON.stringify(content)));
    expect(file.original).toBeUndefined();
    const data = await blobBytes(await exportOffice(file));
    const zip = await JSZip.loadAsync(data);
    const restored = (await readDocx(data)).content;
    expect(new DOMParser().parseFromString(restored.html, 'text/html').body.textContent).toBe(
      Array.from({ length: 70 }, (_, i) => 'line' + String(i + 1).padStart(2, '0')).join(''),
    );
    const section = resolveWordSections(restored)[0];
    expect(section.differentFirstPage).toBe(slot === 'first');
    expect(restored.stories!.evenAndOddHeaders).toBe(slot === 'even');
    for (const key of ['headers', 'footers'] as const)
      for (const otherSlot of ['default', 'first', 'even'] as const) {
        const ref = section[key][otherSlot]!;
        const part = restored.stories!.parts.find((p) =>
          p.relationshipIds.includes(ref.relationshipId),
        )!;
        const html = new DOMParser().parseFromString(part.html, 'text/html');
        expect(html.body.textContent).toBe(
          key === (kind === 'header' ? 'headers' : 'footers') && slot === otherSlot
            ? sample.text
            : '',
        );
        const paragraph = html.querySelector('p')!;
        expect(paragraph.style.fontFamily.replace(/^["']|["']$/g, '')).toBe(sample.story.family);
        expect(paragraph.style.fontSize).toBe(sample.story.size + 'pt');
        expect(paragraph.style.lineHeight).toBe('1');
        expect(paragraph.style.marginBottom).toBe('0pt');
      }
    const styles = wordXml(await zip.file('word/styles.xml')!.async('string'));
    for (const label of ['header', 'footer'])
      expect(
        descendants(styles, 'style').filter((s) => val(child(s, 'name')) === label),
      ).toHaveLength(1);
    expect(Object.keys(zip.files).some((p) => p.endsWith('.odttf'))).toBe(true);
  });

it('keeps fresh blank apply/default options as no-ops, and creation in one body history transaction', () => {
  const { content } = fresh();
  const body = new Editor({
    extensions: [...wordExtensions(), WordStoryState],
    content: content.html,
  });
  try {
    initializeWordStories(body, content);
    const original = body.state.doc.toJSON();
    changeWordStoryPageOptions(body, content, undefined, 'authored-body', {
      differentFirstPage: false,
      evenAndOddHeaders: false,
      headerDistance: 708,
      footerDistance: 708,
    });
    createWordStory(
      body,
      content,
      undefined,
      'authored-body',
      'header',
      'default',
      authoredWordStoryTemplates(content)!.header,
    );
    expect(body.state.doc.toJSON()).toEqual(original);
    expect(body.can().undo()).toBe(false);
    const html = authoredWordStoryTemplates(content)!.header.replace(
      '</p>',
      'Created fresh header</p>',
    );
    createWordStory(body, content, undefined, 'authored-body', 'header', 'default', html);
    const created = body.state.doc.toJSON();
    expect(body.state.doc.attrs.wordStories.parts).toHaveLength(6);
    body.commands.undo();
    expect(body.state.doc.toJSON()).toEqual(original);
    body.commands.redo();
    expect(body.state.doc.toJSON()).toEqual(created);
  } finally {
    body.destroy();
  }
});

it('rejects imported-story provenance without its retained source instead of dropping package dependencies', async () => {
  const { file, content } = fresh();
  content.stories = {
    version: 1,
    evenAndOddHeaders: false,
    parts: [
      {
        path: 'word/header1.xml',
        kind: 'header',
        relationshipIds: ['rHeader'],
        html: '<p>Retained object</p>',
      },
    ],
  };
  await expect(exportOffice(file)).rejects.toThrow('retained DOCX source');
  expect(content.stories.parts[0].html).toBe('<p>Retained object</p>');
});
