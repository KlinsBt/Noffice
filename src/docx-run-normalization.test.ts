import { expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { coalescePlainWordRuns, normalizeWordTextSpace } from './docx-run-normalization';
import { WORD_NS, wordXml } from './docx-import';

const font = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>';
const run = (text: string, properties = font, attributes = '') =>
  `<w:r${attributes}>${properties}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const document = (body: string) =>
  wordXml(`<w:document xmlns:w="${WORD_NS}"><w:body>${body}</w:body></w:document>`);
const serialize = (doc: XMLDocument) => new XMLSerializer().serializeToString(doc);

it('normalizes retained body text metadata without changing text, formatting or run boundaries', () => {
  const doc = document(`<w:p>${run('A\u00bfB A\u00c9B', font, ' w:rsidR="12345678"')}${run('A  B')}${run('\u00a0A\u00a0')}</w:p>`);
  const before = doc.documentElement.textContent;
  const runs = [...doc.getElementsByTagNameNS(WORD_NS, 'r')];
  const properties = runs.map((r) => new XMLSerializer().serializeToString(r.firstElementChild!));
  expect(normalizeWordTextSpace(doc)).toBe(3);
  expect(doc.documentElement.textContent).toBe(before);
  expect([...doc.getElementsByTagNameNS(WORD_NS, 'r')]).toEqual(runs);
  expect(runs[0].getAttributeNS(WORD_NS, 'rsidR')).toBe('12345678');
  expect(runs.map((r) => new XMLSerializer().serializeToString(r.firstElementChild!))).toEqual(properties);
  expect(normalizeWordTextSpace(doc)).toBe(0);
});

it('retains significant whitespace and metadata on opaque or nested content', () => {
  const doc = document(`<w:p>${[' leading', 'trailing ', ' ', 'A&#x9;B', 'A&#xA;B', ''].map((text) => run(text)).join('')}
    <w:r>${font}<w:fldChar w:fldCharType="begin"/><w:t xml:space="preserve">field</w:t></w:r>
    <w:hyperlink w:anchor="keep">${run('linked')}</w:hyperlink>
    <w:r>${font}<w:t xml:space="preserve" data-opaque="keep">opaque</w:t></w:r>
    </w:p><w:tbl><w:tr><w:tc><w:p>${run('cell')}</w:p></w:tc></w:tr></w:tbl>`);
  const before = serialize(doc);
  expect(normalizeWordTextSpace(doc)).toBe(0);
  expect(serialize(doc)).toBe(before);
});

it('matches native merging of the fixture fragment while preserving text and paragraph identity', async () => {
  const zip = await JSZip.loadAsync(readFileSync('tests/fixtures/word-mixed-wrap.docx'));
  const doc = wordXml(await zip.file('word/document.xml')!.async('string'));
  const text = doc.documentElement.textContent;
  const paragraphs = [...doc.getElementsByTagNameNS(WORD_NS, 'p')];
  const ids = paragraphs.map((p) => p.getAttributeNS(WORD_NS, 'rsidR'));
  expect(coalescePlainWordRuns(doc)).toBe(1);
  expect(paragraphs.map((p) => p.getElementsByTagNameNS(WORD_NS, 'r').length)).toEqual([
    1, 13, 13, 1,
  ]);
  expect(doc.documentElement.textContent).toBe(text);
  expect(paragraphs.map((p) => p.getAttributeNS(WORD_NS, 'rsidR'))).toEqual(ids);
  expect(coalescePlainWordRuns(doc)).toBe(0);
});

it.each([
  ['revision attribute', run('a') + run('b', font, ' w:rsidR="12345678"')],
  ['bookmark', run('a') + '<w:bookmarkStart w:id="4" w:name="keep"/>' + run('b')],
  ['comment boundary', run('a') + '<!--keep-->' + run('b')],
  ['unpreserved edge spaces', run('a') + run(' b ').replace(' xml:space="preserve"', '')],
  ['field', run('a') + `<w:r>${font}<w:instrText>PAGE</w:instrText></w:r>` + run('b')],
  [
    'opaque properties',
    run('a', font.replace('</w:rPr>', '<w:lang w:val="de-DE"/></w:rPr>')) +
      run('b', font.replace('</w:rPr>', '<w:lang w:val="de-DE"/></w:rPr>')),
  ],
  ['different font', run('a') + run('b', font.replace('Arial', 'Calibri'))],
  ['hard break', run('a') + `<w:r>${font}<w:br/><w:t>b</w:t></w:r>`],
  ['hyperlink', run('a') + `<w:hyperlink w:anchor="keep">${run('b')}</w:hyperlink>`],
])('preserves the %s boundary verbatim', (_name, body) => {
  const doc = document(`<w:p>${body}</w:p>`);
  const before = serialize(doc);
  expect(coalescePlainWordRuns(doc)).toBe(0);
  expect(serialize(doc)).toBe(before);
});

it('preserves significant spaces and keeps separate paragraphs separate', () => {
  const doc = document(`<w:p>${run(' one ')}${run(' two')}</w:p><w:p>${run('three')}</w:p>`);
  expect(coalescePlainWordRuns(doc)).toBe(1);
  const texts = [...doc.getElementsByTagNameNS(WORD_NS, 't')];
  expect(texts.map((t) => t.textContent)).toEqual([' one  two', 'three']);
  expect(texts[0].getAttributeNS('http://www.w3.org/XML/1998/namespace', 'space')).toBe('preserve');
});

it('does not normalize paragraphs inside opaque drawings or table cells', () => {
  const doc = document(
    `<w:p><w:r><w:drawing><w:txbxContent><w:p>${run('a')}${run('b')}</w:p></w:txbxContent></w:drawing></w:r></w:p><w:tbl><w:tr><w:tc><w:p>${run('a')}${run('b')}</w:p></w:tc></w:tr></w:tbl>`,
  );
  const before = serialize(doc);
  expect(coalescePlainWordRuns(doc)).toBe(0);
  expect(serialize(doc)).toBe(before);
});
