import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { readDocx, wordXml, descendants, val } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { sanitizeWordContent, importFile, nativeBackup, exportOffice } from './formats';
import { newFile } from './model';
import { contentFingerprint } from './office-preservation';
import { hydrateWordStructure, needsWordStructure } from './word-structure';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import reference from '../tests/fixtures/native-word-side-stories.json';

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
const bytes = (name = 'default') =>
  Uint8Array.from(readFileSync(`tests/fixtures/word-side-stories/${name}.docx`)).buffer;
const dom = (html: string) => new DOMParser().parseFromString(html, 'text/html').body;
const blobBytes = (blob: Blob) =>
  new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });
async function setup(name = 'default') {
  const data = bytes(name);
  const read = await readDocx(data);
  const file = newFile('word', name, sanitizeWordContent(read.content));
  file.original = {
    name: `${name}.docx`,
    data,
    contentFingerprint: await contentFingerprint(file.content),
  };
  if (file.content.kind !== 'word') throw Error('Expected Word');
  return { file, read, content: file.content };
}
for (const [name, sample] of Object.entries(reference.cases))
  it(`imports native ${name} stories independently and preserves them through body export`, async () => {
    const { file, read, content } = await setup(name);
    expect(content.stories?.parts).toHaveLength(6);
    expect(dom(content.html).textContent).not.toMatch(/Header|Footer|NOFFICE/);
    expect(content.stories?.evenAndOddHeaders).toBe(sample.native.sections[0].even);
    for (const part of content.stories!.parts) {
      expect(part.html).not.toContain('NOFFICE');
      const section = read.content.docxStructure!.sections[0];
      const ref = [...section.headers, ...section.footers].find((r) =>
        part.relationshipIds.includes(r.relationshipId),
      )!;
      const native = sample.native.sections[0].stories.find(
        (s) =>
          s.kind === (part.kind === 'header' ? 'Headers' : 'Footers') &&
          s.slot === { default: 1, first: 2, even: 3 }[ref.type],
      )!;
      expect(dom(part.html).textContent).toBe((native.text || '').replace(/[\r\v]/g, ''));
      for (const paragraph of dom(part.html).querySelectorAll<HTMLElement>(
        '[data-source-paragraph]',
      )) {
        expect(paragraph.dataset.sourceParagraph).not.toMatch(/^0:/);
        expect(read.paragraphs.get(paragraph.dataset.sourceParagraph!)?.ownerDocument).toBe(
          read.documents.get(part.path),
        );
      }
    }
    const unchanged = await blobBytes(await exportRetainedDocument(file));
    expect(new Uint8Array(unchanged)).toEqual(new Uint8Array(file.original!.data));
    const body = dom(content.html);
    body.querySelector('p')!.append(' edited');
    content.html = body.innerHTML;
    const exported = await blobBytes(await exportRetainedDocument(file));
    const before = await JSZip.loadAsync(file.original!.data),
      after = await JSZip.loadAsync(exported);
    for (const part of content.stories!.parts)
      expect(await after.file(part.path)!.async('uint8array')).toEqual(
        await before.file(part.path)!.async('uint8array'),
      );
    const reopened = await readDocx(exported);
    expect(dom(reopened.content.html).textContent).toContain(' edited');
    expect(sanitizeWordContent(reopened.content).stories).toEqual(content.stories);
  });

it('hydrates old retained snapshots without replacing edited body text or revision', async () => {
  const { file, content } = await setup();
  delete content.stories;
  file.original!.contentFingerprint = await contentFingerprint(content);
  const hydrated = await hydrateWordStructure(file);
  expect(needsWordStructure(hydrated)).toBe(false);
  expect(hydrated.revision).toBe(file.revision);
  expect(hydrated.content.kind === 'word' && hydrated.content.html).toBe(content.html);
  expect(hydrated.original!.contentFingerprint).toBe(await contentFingerprint(hydrated.content));
  content.html += '<p>Existing edit</p>';
  const edited = await hydrateWordStructure(file);
  expect(edited.content.kind === 'word' && edited.content.html).toBe(content.html);
  expect(edited.original!.contentFingerprint).toBe(file.original!.contentFingerprint);
});

it('sanitizes backup stories and rejects unsupported edits without mutating the source', async () => {
  const { file, content } = await setup();
  const original = new Uint8Array(file.original!.data).slice();
  content.stories!.parts[0].html +=
    '<script>bad()</script><img src="https://tracker.invalid/x"><table><tr><td><p onclick="bad()" style="background:url(https://tracker.invalid)">Edited</p></td></tr></table>';
  const input = { name: 'test.noffice', size: 1, text: async () => nativeBackup(file) } as File;
  const restored = await importFile(input);
  if (restored.content.kind !== 'word') throw Error('Expected Word');
  expect(restored.content.stories!.parts[0].html).not.toMatch(/<script\b|javascript:|onclick|tracker|url\(/i);
  expect(restored.content.stories!.parts[0].html).toContain('Edited');
  await expect(exportRetainedDocument(restored)).rejects.toThrow(
    'header/footer structural editing',
  );
  expect(new Uint8Array(file.original!.data)).toEqual(original);
  delete restored.original;
  await expect(exportOffice(restored)).rejects.toThrow('retained DOCX source');
});

it('keeps empty story hydration fingerprint neutral but binds actual story edits', async () => {
  const file = newFile('word');
  if (file.content.kind !== 'word') throw Error('Expected Word');
  const before = await contentFingerprint(file.content);
  file.content.stories = { version: 1, evenAndOddHeaders: false, parts: [] };
  expect(await contentFingerprint(file.content)).toBe(before);
  file.content.stories.evenAndOddHeaders = true;
  expect(await contentFingerprint(file.content)).not.toBe(before);
});

it('exports header/footer text and formatting with history and no unrelated part changes', async () => {
  const { file, content } = await setup('first-even');
  const part = content.stories!.parts.find((p) => dom(p.html).textContent === 'Header first')!;
  expect(part).toBeDefined();
  const originalHtml = part.html;
  const editor = new Editor({ extensions: wordExtensions(), content: originalHtml });
  try {
    editor.commands.setTextSelection({ from: 1, to: editor.state.doc.content.size - 1 });
    editor.commands.insertContent('Edited first header');
    editor.commands.selectAll();
    editor.commands.setBold();
    part.html = editor.getHTML();
    const editedHtml = part.html;
    expect(await contentFingerprint(content)).not.toBe(file.original!.contentFingerprint);
    expect(editor.commands.undo()).toBe(true);
    expect(editor.commands.redo()).toBe(true);
    expect(editor.getHTML()).toBe(editedHtml);
    const exported = await blobBytes(await exportRetainedDocument(file));
    const before = await JSZip.loadAsync(file.original!.data),
      after = await JSZip.loadAsync(exported);
    for (const path of Object.keys(before.files)) {
      if (before.files[path].dir || [part.path, 'word/settings.xml'].includes(path)) continue;
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
    }
    expect(await after.file(part.path)!.async('string')).toContain('Edited first header');
    const reopened = await readDocx(exported);
    const result = reopened.content.stories!.parts.find((p) => p.path === part.path)!;
    expect(dom(result.html).textContent).toBe('Edited first header');
    expect(dom(result.html).querySelector('strong')?.textContent).toBe('Edited first header');
    expect(reopened.content.html).toBe(content.html);
    part.html = originalHtml;
    expect(new Uint8Array(await blobBytes(await exportRetainedDocument(file)))).toEqual(
      new Uint8Array(file.original!.data),
    );
  } finally {
    editor.destroy();
  }
});

it('splits and joins retained story paragraphs with stable parts and independent paragraph identities', async () => {
  const { file, content } = await setup();
  const part = content.stories!.parts.find((p) => dom(p.html).textContent === 'Header default')!;
  const originalHtml = part.html;
  const editor = new Editor({ extensions: wordExtensions(), content: part.html });
  try {
    editor.commands.setTextSelection(8);
    expect(editor.commands.splitBlock()).toBe(true);
    part.html = editor.getHTML();
    expect(dom(part.html).querySelectorAll('p')).toHaveLength(2);
    const exported = await blobBytes(await exportRetainedDocument(file));
    const read = await readDocx(exported);
    const split = read.content.stories!.parts.find((p) => p.path === part.path)!;
    const paragraphs = [...dom(split.html).querySelectorAll('p')];
    expect(paragraphs.map((p) => p.textContent)).toEqual(['Header ', 'default']);
    expect(new Set(paragraphs.map((p) => p.dataset.sourceParagraph)).size).toBe(2);
    const before = await JSZip.loadAsync(file.original!.data),
      after = await JSZip.loadAsync(exported);
    for (const path of Object.keys(before.files)) {
      if (before.files[path].dir || [part.path, 'word/settings.xml'].includes(path)) continue;
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
    }
    // Join against the newly retained two-paragraph source, not the old model.
    file.original!.data = exported;
    file.content = sanitizeWordContent(read.content);
    if (file.content.kind !== 'word') throw Error('Expected Word');
    const joined = file.content.stories!.parts.find((p) => p.path === part.path)!;
    editor.commands.setContent(joined.html, { parseOptions: { preserveWhitespace: true } });
    editor.commands.setTextSelection(editor.state.doc.firstChild!.nodeSize + 1);
    expect(editor.commands.joinBackward()).toBe(true);
    joined.html = editor.getHTML();
    const reopened = await readDocx(await blobBytes(await exportRetainedDocument(file)));
    expect(
      dom(reopened.content.stories!.parts.find((p) => p.path === part.path)!.html).textContent,
    ).toBe('Header default');
    expect(
      dom(reopened.content.stories!.parts.find((p) => p.path === part.path)!.html).querySelectorAll(
        'p',
      ),
    ).toHaveLength(1);
    expect(originalHtml).toContain('Header default');
  } finally {
    editor.destroy();
  }
});

it('preserves the source run session when joining a story and distinguishes retyped identical text', async () => {
  const { file, content } = await setup();
  const part = content.stories!.parts.find((p) => dom(p.html).textContent === 'Header default')!;
  const first = new Editor({
    extensions: wordExtensions({ retainedEditRuns: true }),
    content: part.html,
  });
  try {
    first.commands.setTextSelection(first.state.doc.content.size - 1);
    first.commands.splitBlock();
    first.commands.insertContent('extra');
    part.html = first.getHTML();
    expect(part.html).toContain('data-word-edit-run');
    const exported = await blobBytes(await exportRetainedDocument(file));
    const read = await readDocx(exported);
    file.original!.data = exported;
    file.content = sanitizeWordContent(read.content);
    if (file.content.kind !== 'word') throw Error('Expected Word');
    const retained = file.content.stories!.parts.find((p) => p.path === part.path)!;
    const baseline = retained.html;
    const xml = wordXml(await (await JSZip.loadAsync(exported)).file(part.path)!.async('string'));
    const sourceSession = val(
      descendants(xml, 'p')[1].getElementsByTagNameNS(xml.documentElement.namespaceURI!, 'r')[0],
      'rsidR',
    );
    expect(sourceSession).toMatch(/^[a-f\d]{8}$/i);
    const editor = new Editor({
      extensions: wordExtensions({ retainedEditRuns: true }),
      content: baseline,
      parseOptions: { preserveWhitespace: true },
    });
    try {
      editor.commands.setTextSelection(editor.state.doc.firstChild!.nodeSize + 1);
      expect(editor.commands.joinBackward()).toBe(true);
      retained.html = editor.getHTML();
      const joined = wordXml(
        await (
          await JSZip.loadAsync(await blobBytes(await exportRetainedDocument(file)))
        )
          .file(part.path)!
          .async('string'),
      );
      const runs = [...descendants(joined, 'p')[0].children].filter((e) => e.localName === 'r');
      expect(
        runs.map((r) =>
          descendants(r, 't')
            .map((t) => t.textContent)
            .join(''),
        ),
      ).toEqual(['Header default', 'extra']);
      expect(val(runs[1], 'rsidR')).toBe(sourceSession);
      editor.commands.setContent(baseline, { parseOptions: { preserveWhitespace: true } });
      editor.commands.setTextSelection(editor.state.doc.firstChild!.nodeSize - 1);
      editor.commands.insertContent('extra');
      editor.commands.deleteRange({
        from: editor.state.doc.firstChild!.nodeSize,
        to: editor.state.doc.content.size,
      });
      retained.html = editor.getHTML();
      expect(editor.getText()).toBe('Header defaultextra');
      const retyped = wordXml(
        await (
          await JSZip.loadAsync(await blobBytes(await exportRetainedDocument(file)))
        )
          .file(part.path)!
          .async('string'),
      );
      const newRuns = [...descendants(retyped, 'p')[0].children].filter((e) => e.localName === 'r');
      expect(
        newRuns.map((r) =>
          descendants(r, 't')
            .map((t) => t.textContent)
            .join(''),
        ),
      ).toEqual(['Header default', 'extra']);
      expect(val(newRuns[1], 'rsidR')).not.toBe(sourceSession);
    } finally {
      editor.destroy();
    }
  } finally {
    first.destroy();
  }
});

it('rejects foreign story identities and deletion of a hidden bookmark without changing originals', async () => {
  const { file, content } = await setup();
  const part = content.stories!.parts.find((p) => dom(p.html).textContent === 'Header default')!;
  const original = part.html;
  part.html = original.replace(/data-source-paragraph="[^"]+"/, 'data-source-paragraph="0:0"');
  await expect(exportRetainedDocument(file)).rejects.toThrow('paragraph identities');
  const zip = await JSZip.loadAsync(file.original!.data);
  const xml = await zip.file(part.path)!.async('string');
  zip.file(
    part.path,
    xml.replace(
      '</w:p>',
      '<w:bookmarkStart w:id="900" w:name="KeepStoryAnchor"/><w:bookmarkEnd w:id="900"/></w:p>',
    ),
  );
  const data = await zip.generateAsync({ type: 'arraybuffer' });
  const read = await readDocx(data);
  file.original!.data = data;
  file.content = sanitizeWordContent(read.content);
  if (file.content.kind !== 'word') throw Error('Expected Word');
  file.content.stories!.parts.find((p) => p.path === part.path)!.html = '<p>Replacement</p>';
  await expect(exportRetainedDocument(file)).rejects.toThrow('containing document structures');
  expect(file.original!.data).toBe(data);
});

it('rebases a story hyperlink without colliding with an identically named body relationship', async () => {
  const zip = await JSZip.loadAsync(bytes());
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const baseline = await readDocx(bytes());
  const path = baseline.content.stories!.parts.find(
    (p) => dom(p.html).textContent === 'Header default',
  )!.path;
  const header = await zip.file(path)!.async('string');
  const xml = wordXml(header),
    run = descendants(xml, 'r')[0];
  const link = xml.createElementNS(run.namespaceURI, 'w:hyperlink');
  link.setAttributeNS(R, 'r:id', 'storyLink');
  run.replaceWith(link);
  link.append(run);
  zip.file(path, new XMLSerializer().serializeToString(xml));
  zip.file(
    `word/_rels/${path.slice(5)}.rels`,
    `<Relationships xmlns="${REL}"><Relationship Id="storyLink" Type="${R}/hyperlink" Target="https://header.example/" TargetMode="External"/></Relationships>`,
  );
  const main = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file(
    'word/_rels/document.xml.rels',
    main.replace(
      '</Relationships>',
      `<Relationship Id="storyLink" Type="${R}/hyperlink" Target="https://body.example/" TargetMode="External"/></Relationships>`,
    ),
  );
  const read = await readDocx(await zip.generateAsync({ type: 'arraybuffer' }));
  expect(
    dom(read.content.stories!.parts.find((p) => p.path === path)!.html)
      .querySelector('a')
      ?.getAttribute('href'),
  ).toBe('https://header.example/');
  expect(read.content.html).not.toContain('header.example');
  expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain(
    'https://body.example/',
  );
});

it('adds blank-story style templates to legacy snapshots without replacing story edits', async () => {
  const { file, content } = await setup();
  const templates = structuredClone(content.stories!.emptyTemplates);
  delete content.stories!.emptyTemplates;
  file.original!.contentFingerprint = await contentFingerprint(content);
  const hydrated = await hydrateWordStructure(file);
  if (hydrated.content.kind !== 'word') throw Error('Expected Word');
  expect(hydrated.content.stories!.emptyTemplates).toEqual(templates);
  expect(hydrated.original!.contentFingerprint).toBe(await contentFingerprint(hydrated.content));
  content.stories!.parts[0].html += '<p>Existing story edit</p>';
  const edited = await hydrateWordStructure(file);
  if (edited.content.kind !== 'word') throw Error('Expected Word');
  expect(edited.content.stories!.parts[0].html).toBe(content.stories!.parts[0].html);
  expect(edited.original!.contentFingerprint).toBe(file.original!.contentFingerprint);
  expect(edited.revision).toBe(file.revision);
  expect(await hydrateWordStructure(edited)).toBe(edited);
});
