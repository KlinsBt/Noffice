import { expect, it } from 'vitest';
import { splitWordColumnRuns } from './docx-column-runs';
import { descendants, WORD_NS, wordXml } from './docx-import';

const create = (body: string) =>
  wordXml(`<w:document xmlns:w="${WORD_NS}"><w:body><w:p>${body}</w:p></w:body></w:document>`);
const run = (text: string) =>
  `<w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>${text}</w:t></w:r>`;
const layout = (text: string, boundaries: number[]) => ({
  content: '',
  paragraphs: [{ text, boundaries }],
});

it('splits current column offsets while retaining font properties, text and source metadata', () => {
  const doc = create('<w:bookmarkStart w:id="1" w:name="keep"/>' + run('one two three four'));
  const p = descendants(doc, 'p')[0];
  p.setAttributeNS(WORD_NS, 'w:rsidR', '00000001');
  splitWordColumnRuns(doc, layout('one two three four', [4, 14]));
  expect(descendants(doc, 't').map((t) => t.textContent)).toEqual(['one ', 'two three ', 'four']);
  expect(descendants(doc, 'rFonts')).toHaveLength(3);
  expect(descendants(doc, 'bookmarkStart')).toHaveLength(1);
  expect(p.getAttributeNS(WORD_NS, 'rsidR')).toBe('00000001');
  expect(descendants(doc, 'br')).toHaveLength(0);
  expect(descendants(doc, 'lastRenderedPageBreak')).toHaveLength(0);
  const xml = new XMLSerializer().serializeToString(doc);
  splitWordColumnRuns(doc, layout('one two three four', [4, 14]));
  expect(new XMLSerializer().serializeToString(doc)).toBe(xml);
});

it('counts line controls and preserves them when a run crosses a column boundary', () => {
  const doc = create('<w:r><w:rPr><w:b/></w:rPr><w:t>one</w:t><w:br/><w:t>two three</w:t></w:r>');
  splitWordColumnRuns(doc, layout('one\ntwo three', [8]));
  expect(descendants(doc, 'r').map((r) => r.textContent)).toEqual(['onetwo ', 'three']);
  expect(descendants(doc, 'br')).toHaveLength(1);
  expect(descendants(doc, 'b')).toHaveLength(2);
});

it('keeps optional controls and literal hyphens distinct when a column splits their run', () => {
  const doc = create('<w:r><w:rPr><w:b/></w:rPr><w:t>ab</w:t><w:softHyphen/><w:t>c\u00add</w:t></w:r>');
  splitWordColumnRuns(doc, layout('ab\u001fc\u00add', [2, 4]));
  expect(descendants(doc, 'r').map(r => r.textContent)).toEqual(['ab', 'c', '\u00add']);
  expect(descendants(doc, 'r')[1].children[1].localName).toBe('softHyphen');
  expect(descendants(doc, 'softHyphen')).toHaveLength(1);
  expect(descendants(doc, 'b')).toHaveLength(3);
});

it.each(['<w:softHyphen w:unknown="preserve"/>', '<w:softHyphen>hidden</w:softHyphen>'])
  ('rejects malformed optional controls before splitting their source run: %s', control => {
    const doc = create(`<w:r><w:t>ab</w:t>${control}<w:t>cd</w:t></w:r>`);
    const before = new XMLSerializer().serializeToString(doc);
    expect(() => splitWordColumnRuns(doc, layout('ab\u001fcd', [2]))).toThrow('unsupported');
    expect(new XMLSerializer().serializeToString(doc)).toBe(before);
  });

it.each([
  ['different text', 'changed', [2]],
  ['out of bounds', 'one two', [8]],
  ['reversed offsets', 'one two', [4, 2]],
  ['fractional offset', 'one two', [2.5]],
])('rejects %s before writing any column runs', (_name, text, boundaries) => {
  const doc = create(run('one two'));
  const before = new XMLSerializer().serializeToString(doc);
  expect(() => splitWordColumnRuns(doc, layout(text as string, boundaries as number[]))).toThrow();
  expect(new XMLSerializer().serializeToString(doc)).toBe(before);
});

it('rejects a split inside a surrogate pair or opaque hyperlink', () => {
  expect(() => splitWordColumnRuns(create(run('a😀b')), layout('a😀b', [2]))).toThrow();
  expect(() =>
    splitWordColumnRuns(
      create(`<w:hyperlink w:anchor="keep">${run('one two')}</w:hyperlink>`),
      layout('one two', [4]),
    ),
  ).toThrow('unsupported');
});
