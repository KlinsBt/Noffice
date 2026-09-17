import { expect, it, vi } from 'vitest';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  ExternalHyperlink,
  FootnoteReferenceRun,
} from 'docx';
import JSZip from 'jszip';
import { newFile } from './model';
import { readDocx, descendants, wordXml, val } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { exportOffice } from './formats';

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
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const bytes = (blob: Blob) =>
  new Promise<ArrayBuffer>((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.readAsArrayBuffer(blob);
  });
async function fixture() {
  const link = (text: string) =>
    new ExternalHyperlink({
      link: 'https://example.com/original',
      children: [
        new TextRun({ text, bold: true }),
        new TextRun({ text: ' detail', italics: true }),
      ],
    });
  const data = await Packer.toArrayBuffer(
    new Document({
      footnotes: { 2: { children: [new Paragraph({ children: [link('Note')] })] } },
      sections: [
        {
          children: [
            new Paragraph({
              children: [new TextRun('Before '), link('Alpha'), new TextRun(' after')],
            }),
            new Paragraph({ children: [link('Other')] }),
            new Paragraph({ children: [new TextRun('Plain text'), new FootnoteReferenceRun(2)] }),
            new Paragraph('New destination'),
          ],
        },
      ],
    }),
  );
  const zip = await JSZip.loadAsync(data);
  const doc = wordXml(await zip.file('word/document.xml')!.async('string'));
  descendants(doc, 'hyperlink')[0].setAttributeNS(
    doc.documentElement.namespaceURI,
    'w:tooltip',
    'Retain this tip',
  );
  zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
  zip.file('customXml/keep.xml', '<keep>unchanged</keep>');
  const input = await zip.generateAsync({ type: 'arraybuffer' }),
    read = await readDocx(input),
    file = newFile('word', 'Links', read.content);
  file.original = { name: 'Links.docx', data: input };
  const html = new DOMParser().parseFromString(read.content.html, 'text/html');
  function save() {
    if (file.content.kind === 'word') file.content.html = html.body.innerHTML;
  }
  async function output() {
    save();
    const data = await bytes(await exportRetainedDocument(file));
    return { data, zip: await JSZip.loadAsync(data) };
  }
  return { file, html, input, output };
}
it('formats existing hyperlink runs while preserving destinations, wrapper metadata and unrelated parts', async () => {
  const f = await fixture(),
    first = f.html.querySelector('a')!;
  first.innerHTML = `<u>${first.innerHTML}</u>`;
  const result = await f.output(),
    doc = wordXml(await result.zip.file('word/document.xml')!.async('string'));
  expect(val(descendants(doc, 'hyperlink')[0], 'tooltip')).toBe('Retain this tip');
  expect(descendants(descendants(doc, 'hyperlink')[0], 'u').map((e) => val(e))).toEqual([
    'single',
    'single',
  ]);
  expect(descendants(descendants(doc, 'hyperlink')[0], 'b').length).toBeGreaterThan(0);
  const before = await JSZip.loadAsync(f.input);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'word/document.xml')
      expect(await result.zip.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
});
it('changes and removes links without retargeting another link that used the same destination', async () => {
  const f = await fixture();
  f.html.querySelector('a')!.setAttribute('href', 'https://example.com/revised?a=1&b=2');
  const result = await f.output(),
    doc = wordXml(await result.zip.file('word/document.xml')!.async('string'));
  const rels = wordXml(await result.zip.file('word/_rels/document.xml.rels')!.async('string'));
  const target = (link: Element) =>
    Array.from(rels.documentElement.children)
      .find((e) => e.getAttribute('Id') === link.getAttributeNS(R, 'id'))
      ?.getAttribute('Target');
  expect(target(descendants(doc, 'hyperlink')[0])).toBe('https://example.com/revised?a=1&b=2');
  expect(target(descendants(doc, 'hyperlink')[1])).toBe('https://example.com/original');
  const first = f.html.querySelector('a')!;
  first.replaceWith(...first.childNodes);
  const removed = await f.output();
  expect(
    descendants(wordXml(await removed.zip.file('word/document.xml')!.async('string')), 'hyperlink'),
  ).toHaveLength(1);
  expect((await readDocx(removed.data)).content.html).toContain('Alpha');
});
it('creates part-local relationships for new body and footnote links and leaves source relationships intact', async () => {
  const f = await fixture();
  const note = Array.from(f.html.querySelectorAll('a')).find((a) =>
    a.textContent?.startsWith('Note'),
  )!;
  note.href = 'mailto:notes@example.com';
  const p = Array.from(f.html.querySelectorAll('p')).find(
    (p) => p.textContent === 'New destination',
  )!;
  p.innerHTML = '<a href="mailto:notes@example.com"><strong>New</strong> destination</a>';
  const result = await f.output();
  for (const path of ['word/_rels/document.xml.rels', 'word/_rels/footnotes.xml.rels']) {
    const rels = wordXml(await result.zip.file(path)!.async('string'));
    expect(
      Array.from(rels.documentElement.children).some(
        (e) =>
          e.getAttribute('Target') === 'mailto:notes@example.com' &&
          e.getAttribute('TargetMode') === 'External',
      ),
    ).toBe(true);
  }
  const reimport = await readDocx(result.data);
  const footnotes = wordXml(await result.zip.file('word/footnotes.xml')!.async('string'));
  expect(
    descendants(
      descendants(footnotes, 'footnote').find((e) => val(e, 'id') === '2')!,
      'footnoteRef',
    ),
  ).toHaveLength(1);
  expect(reimport.content.html.match(/mailto:notes@example.com/g)).toHaveLength(2);
});
it('writes mixed font/emphasis inside hyperlinks in newly created DOCX files', async () => {
  const file = newFile('word', 'New links');
  if (file.content.kind !== 'word') throw Error();
  file.content.html =
    '<p><a href="https://example.com"><strong>Bold</strong><em> italic</em><span style="font-family:Georgia;font-size:18pt"> Georgia</span></a></p>';
  const exported = await exportOffice(file),
    zip = await JSZip.loadAsync(await bytes(exported)),
    doc = wordXml(await zip.file('word/document.xml')!.async('string'));
  const link = descendants(doc, 'hyperlink')[0];
  expect(
    descendants(link, 't')
      .map((t) => t.textContent)
      .join(''),
  ).toBe('Bold italic Georgia');
  expect(descendants(link, 'b').length).toBeGreaterThan(0);
  expect(descendants(link, 'i').length).toBeGreaterThan(0);
  expect(descendants(link, 'rFonts').some((e) => val(e, 'ascii') === 'Georgia')).toBe(true);
});

it('adds a linked paragraph to a retained document without changing its existing content', async () => {
  const f = await fixture();
  const p = f.html.createElement('p');
  p.innerHTML = '<a href="https://example.com/added"><strong>Added link</strong></a>';
  f.html.querySelector('p')!.after(p);
  const result = await f.output(),
    doc = wordXml(await result.zip.file('word/document.xml')!.async('string'));
  expect(
    descendants(doc, 'hyperlink').map((e) =>
      descendants(e, 't')
        .map((t) => t.textContent)
        .join(''),
    ),
  ).toEqual(['Alpha detail', 'Added link', 'Other detail']);
  expect((await readDocx(result.data)).content.html).toContain('https://example.com/added');
});
