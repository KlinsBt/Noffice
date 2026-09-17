import { expect, it, vi } from 'vitest';
import { Editor } from '@tiptap/core';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { wordExtensions } from './word-extensions';
import { wordKerningValue, wordKerningStyle } from './word-kerning';
import { readDocx, WORD_NS } from './docx-import';
import { sanitizeHTML, exportOffice } from './formats';
import { newFile } from './model';
import { hydrateParagraphKerning, hydrateParagraphMarkKerning } from './word-kerning-migration';
import { emptyWordStory } from './word-story-create';
import { hydrateWordStructure } from './word-structure';
import { wordJSON } from './word-extensions';
import { contentFingerprint } from './office-preservation';
import reference from '../tests/fixtures/native-word-kerning.json';

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

it('retains native zero and threshold/inherited kerning without leaking reading-copy styles', async () => {
  const { content, messages } = await readDocx(
    Uint8Array.from(readFileSync('tests/fixtures/word-kerning.docx')).buffer,
  );
  const editor = new Editor({ extensions: wordExtensions(), content: sanitizeHTML(content.html) });
  try {
    const paragraphs = [...editor.view.dom.querySelectorAll('p')];
    expect(paragraphs).toHaveLength(12);
    for (const [i, config] of reference.configuration.entries()) {
      const span = paragraphs[i].querySelector<HTMLElement>('span')!;
      expect(span.getAttribute('data-word-kerning')).toBe(
        // Word removes a standalone zero on save when there is no inherited kern.
        config.threshold === null || config.threshold === 0 ? null : String(config.threshold),
      );
      if (config.threshold !== null && config.threshold !== 0)
        expect(span.style.fontKerning).toBe(
          config.threshold > 0 && config.size * 2 >= config.threshold ? 'normal' : 'none',
        );
    }
    expect(editor.getHTML()).not.toContain('NOFFICE');
    expect(messages.filter((m) => m.message.includes('KERN'))).toEqual([]);
  } finally {
    editor.destroy();
  }
});

it('changes kerning at its retained size threshold with atomic history and sanitized reload', () => {
  const editor = new Editor({
    extensions: wordExtensions(),
    content: '<p><span data-word-kerning="24" style="font-size:10pt">11 AV To</span></p>',
  });
  try {
    editor.commands.selectAll();
    editor.commands.setFontSize('20pt');
    expect(editor.view.dom.querySelector('span')!.style.fontKerning).toBe('normal');
    expect(editor.commands.undo()).toBe(true);
    expect(editor.view.dom.querySelector('span')!.style.fontKerning).toBe('none');
    expect(editor.commands.redo()).toBe(true);
    const restored = new Editor({
      extensions: wordExtensions(),
      content: sanitizeHTML(editor.getHTML()),
    });
    try {
      expect(restored.getJSON()).toEqual(editor.getJSON());
    } finally {
      restored.destroy();
    }
  } finally {
    editor.destroy();
  }
});

it('writes kerning thresholds and explicit zero to newly authored DOCX runs', async () => {
  const file = newFile('word');
  if (file.content.kind !== 'word') throw Error('Wrong fixture kind');
  file.content.html =
    '<p><span data-word-kerning="24" style="font-size:20pt">Active</span><span data-word-kerning="0" style="font-size:20pt">Disabled</span></p>';
  const bytes = await exportOffice(file);
  const zip = await JSZip.loadAsync(bytes);
  const xml = new DOMParser().parseFromString(
    await zip.file('word/document.xml')!.async('string'),
    'application/xml',
  );
  expect(
    [...xml.getElementsByTagNameNS(WORD_NS, 'kern')].map((node) =>
      node.getAttributeNS(WORD_NS, 'val'),
    ),
  ).toEqual(['24', '0pt']);
  const imported = await readDocx(await zip.generateAsync({ type: 'arraybuffer' }));
  expect(imported.content.html).toContain('data-word-kerning="0"');
  expect(imported.content.html).toContain('data-word-kerning="24"');
});

it('rejects malformed/unbounded metadata and keeps native zero disabled', () => {
  for (const value of [-1, 3277, Infinity, '12.5', 'url(x)', {}, ''])
    expect(wordKerningValue(value)).toBeNull();
  expect(wordKerningStyle(0, '20pt')).toBe('none');
  expect(wordKerningStyle(24, '16px')).toBe('normal');
  const editor = new Editor({
    extensions: wordExtensions(),
    content: '<p><span data-word-kerning="99999999">Unsafe metadata</span></p>',
  });
  try {
    expect(editor.getHTML()).not.toContain('data-word-kerning');
  } finally {
    editor.destroy();
  }
});

it('hydrates legacy snapshots without losing text edits, revision, originals or unchanged identity', async () => {
  const data = Uint8Array.from(readFileSync('tests/fixtures/word-kerning.docx')).buffer;
  const { content } = await readDocx(data);
  const expected = wordJSON(content.html);
  const file = newFile('word', 'Legacy kerning', content);
  if (file.content.kind !== 'word') throw Error('Wrong fixture kind');
  delete file.content.kerningVersion;
  const dom = new DOMParser().parseFromString(file.content.html, 'text/html');
  for (const span of dom.querySelectorAll<HTMLElement>('[data-word-kerning]')) {
    span.removeAttribute('data-word-kerning');
    span.style.removeProperty('font-kerning');
  }
  file.content.html = dom.body.innerHTML;
  file.original = {
    name: 'kerning.docx',
    data,
    contentFingerprint: await contentFingerprint(file.content),
  };
  const before = structuredClone(file);
  const restored = await hydrateWordStructure(file);
  if (restored.content.kind !== 'word') throw Error('Wrong fixture kind');
  expect(wordJSON(restored.content.html)).toEqual(expected);
  expect(restored.original!.data).toBe(data);
  expect(restored.original!.contentFingerprint).toBe(await contentFingerprint(restored.content));
  expect(await hydrateWordStructure(restored)).toBe(restored);
  dom.querySelectorAll('p')[2].append('!');
  file.content.html = dom.body.innerHTML;
  const edited = await hydrateWordStructure(file);
  if (edited.content.kind !== 'word') throw Error('Wrong fixture kind');
  expect(edited.content.html).toContain('!');
  expect(edited.original!.contentFingerprint).toBe(before.original!.contentFingerprint);
  expect(edited.revision).toBe(before.revision);
  expect(file.content.kerningVersion).toBeUndefined();
});

it('recovers unedited mixed kerning across changed span boundaries and rejects ambiguous edited text', () => {
  const paragraph = (html: string) =>
    new DOMParser()
      .parseFromString(`<p style="font-size:10pt">${html}</p>`, 'text/html')
      .querySelector('p')!;
  const source = paragraph(
    '<span data-word-kerning="16">AB</span><span data-word-kerning="0">CD</span>',
  );
  const current = paragraph('<strong>A</strong>BC<em>D</em>');
  hydrateParagraphKerning(current, source);
  expect(
    [...current.querySelectorAll('[data-word-kerning]')].map((span) => [
      span.textContent,
      span.getAttribute('data-word-kerning'),
    ]),
  ).toEqual([
    ['A', '16'],
    ['B', '16'],
    ['C', '0'],
    ['D', '0'],
  ]);
  const changed = paragraph('ACBD');
  const before = changed.outerHTML;
  expect(() => hydrateParagraphKerning(changed, source)).toThrow('native backup');
  expect(changed.outerHTML).toBe(before);
});

it('restores empty paragraph mark inheritance without replacing explicit zero or nonempty run provenance', () => {
  const p = (html: string) =>
    new DOMParser().parseFromString(html, 'text/html').querySelector('p')!;
  const original = p('<p data-word-paragraph-kerning="2" style="font-size:11pt"></p>');
  const current = p('<p><strong>AV</strong><span data-word-kerning="0"> To</span></p>');
  hydrateParagraphMarkKerning(current, original);
  expect(current.dataset.wordParagraphKerning).toBe('2');
  expect(current.querySelector('strong span')?.getAttribute('data-word-kerning')).toBe('2');
  expect(current.querySelector('[data-word-kerning="0"]')?.textContent).toBe(' To');
  expect(current.textContent).toBe('AV To');
  const once = current.outerHTML;
  hydrateParagraphMarkKerning(current, original);
  expect(current.outerHTML).toBe(once);
  const nonempty = p('<p data-word-paragraph-kerning="24">Old source</p>');
  const edited = p('<p data-word-paragraph-kerning="0">Edited</p>');
  hydrateParagraphMarkKerning(edited, nonempty);
  expect(edited.outerHTML).toBe('<p data-word-paragraph-kerning="0">Edited</p>');
});

it('migrates old generated empty stories without discarding edits, explicit overrides or original identity', async () => {
  const data = Uint8Array.from(
    readFileSync('tests/fixtures/word-story-create-existing-styles.docx'),
  ).buffer;
  const { content } = await readDocx(data);
  const part = emptyWordStory(content.stories!, 'header');
  const dom = new DOMParser().parseFromString(part.html, 'text/html');
  const paragraph = dom.querySelector('p')!;
  expect(paragraph.dataset.wordParagraphKerning).toBe('2');
  paragraph.removeAttribute('data-word-paragraph-kerning');
  paragraph.innerHTML =
    '<span style="font-family:Calibri;font-size:11pt">Saved text</span><span data-word-kerning="0" style="font-size:11pt">AV</span>';
  part.html = dom.body.innerHTML;
  content.stories!.parts.push(part);
  delete content.paragraphKerningVersion;
  const file = newFile('word', 'Legacy generated story', content);
  file.original = { name: 'legacy.docx', data, contentFingerprint: 'edited-original' };
  const before = structuredClone(file);
  const migrated = await hydrateWordStructure(file);
  if (migrated.content.kind !== 'word') throw Error('Expected Word');
  const restored = new DOMParser().parseFromString(
    migrated.content.stories!.parts.find((p) => p.path === part.path)!.html,
    'text/html',
  );
  expect(restored.querySelector('p')!.dataset.wordParagraphKerning).toBe('2');
  expect(restored.querySelector('[data-word-kerning="2"]')!.textContent).toBe('Saved text');
  expect(restored.querySelector('[data-word-kerning="0"]')!.textContent).toBe('AV');
  expect(migrated.original).toBe(file.original);
  expect(migrated.revision).toBe(file.revision);
  expect(await hydrateWordStructure(migrated)).toBe(migrated);
  expect(file).toEqual(before);
});
