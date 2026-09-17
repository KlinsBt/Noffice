import { it, expect } from 'vitest';
import { prepareLiveSections } from './docx-live-sections';
import { readDocxStructure } from './docx-sections';
import { wordXml, WORD_NS, descendants } from './docx-import';

it('materializes all inherited slots after deletion in OOXML element order', () => {
  const refs = ['header', 'footer']
    .flatMap((kind) =>
      ['default', 'first', 'even'].map(
        (type) => `<w:${kind}Reference w:type="${type}" r:id="${kind}-${type}"/>`,
      ),
    )
    .join('');
  const doc = wordXml(
    `<w:document xmlns:w="${WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:pPr><w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p><w:p/><w:sectPr><w:pgSz w:w="16838" w:h="11906"/></w:sectPr></w:body></w:document>`,
  );
  const source = readDocxStructure(doc);
  const commit = prepareLiveSections(descendants(doc, 'body')[0], source, {
    version: 1,
    finalSectionId: source.sections[1].id,
    breaks: [],
  });
  expect(commit([])).toBe(true);
  const final = descendants(doc, 'sectPr')[0];
  expect([...final.children].map((e) => e.localName)).toEqual([
    'headerReference',
    'headerReference',
    'headerReference',
    'footerReference',
    'footerReference',
    'footerReference',
    'pgSz',
  ]);
  expect(source.sections[1].headers).toEqual([]);
});
it('preserves the original section and tracked-property order when the boundary did not move', () => {
  const doc = wordXml(
    `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:pPr><w:sectPr/><w:pPrChange w:id="1"><w:pPr/></w:pPrChange></w:pPr></w:p><w:p/><w:sectPr/></w:body></w:document>`,
  );
  const source = readDocxStructure(doc),
    before = new XMLSerializer().serializeToString(doc);
  const commit = prepareLiveSections(descendants(doc, 'body')[0], source, {
    version: 1,
    finalSectionId: source.sections[1].id,
    breaks: [{ sectionId: source.sections[0].id, paragraph: 0 }],
  });
  expect(commit(descendants(doc, 'p'))).toBe(false);
  expect(new XMLSerializer().serializeToString(doc)).toBe(before);
});
it('rejects a live break targeting an unsupported output paragraph', () => {
  const doc = wordXml(
    `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:pPr><w:sectPr/></w:pPr></w:p><w:p/><w:sectPr/></w:body></w:document>`,
  );
  const source = readDocxStructure(doc);
  const commit = prepareLiveSections(descendants(doc, 'body')[0], source, {
    version: 1,
    finalSectionId: source.sections[1].id,
    breaks: [{ sectionId: source.sections[0].id, paragraph: 1 }],
  });
  expect(() => commit([])).toThrow('no supported output paragraph');
});
