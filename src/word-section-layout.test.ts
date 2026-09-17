import { expect, it } from 'vitest';
import { readDocxStructure } from './docx-sections';
import { wordXml, WORD_NS } from './docx-import';
import { resolveWordSections, singleSectionPage } from './word-section-layout';
import type { WordContent } from './model';

function content(sections: string[]): WordContent {
  const main = wordXml(
    `<w:document xmlns:w="${WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${sections.map((s, i) => (i === sections.length - 1 ? `<w:p/><w:sectPr>${s}</w:sectPr>` : `<w:p><w:pPr><w:sectPr>${s}</w:sectPr></w:pPr></w:p>`)).join('')}</w:body></w:document>`,
  );
  return {
    kind: 'word',
    html: '<p/>',
    paper: 'a4',
    margin: 'normal',
    orientation: 'landscape',
    docxStructure: readDocxStructure(main),
  };
}
const letter =
  '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1800" w:top="900" w:bottom="1000" w:header="400" w:footer="500" w:gutter="0"/>';
const a4 =
  '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/>';

it('uses the following section geometry after a live boundary is deleted', () => {
  const c = content([letter, a4]);
  c.sectionState = { version: 1, finalSectionId: c.docxStructure!.sections[1].id, breaks: [] };
  expect(resolveWordSections(c)).toHaveLength(1);
  expect(singleSectionPage(c)).toMatchObject({ width: 16838, height: 11906 });
  expect(c.docxStructure!.sections).toHaveLength(2);
});

it('reports a materialized header as direct while retaining its source provenance', () => {
  const c = content(['<w:headerReference w:type="default" r:id="header1"/>' + letter, a4]);
  const sourceId = c.docxStructure!.sections[0].id;
  expect(resolveWordSections(c)[1].headers.default).toMatchObject({ inherited: true });
  c.sectionState = { version: 1, finalSectionId: c.docxStructure!.sections[1].id, breaks: [] };
  expect(resolveWordSections(c)[0].headers.default).toEqual({
    relationshipId: 'header1',
    sourceSectionId: sourceId,
    inherited: false,
  });
});

it('retains distinct exact geometry and paragraph identities without changing the source', () => {
  const c = content([letter, a4]),
    original = JSON.stringify(c);
  const result = resolveWordSections(c);
  expect(result.map((s) => [s.width, s.height, s.orientation])).toEqual([
    [12240, 15840, 'portrait'],
    [16838, 11906, 'landscape'],
  ]);
  expect(result[0].margins).toEqual({
    left: 1440,
    right: 1800,
    top: 900,
    bottom: 1000,
    header: 400,
    footer: 500,
    gutter: 0,
  });
  expect(result.map((s) => s.paragraphs)).toEqual([['0:0'], ['0:1']]);
  expect(JSON.stringify(c)).toBe(original);
});
it('inherits each header/footer slot independently across several sections', () => {
  const c = content([
    letter +
      '<w:headerReference w:type="default" r:id="odd"/><w:headerReference w:type="even" r:id="even"/><w:footerReference w:type="first" r:id="foot"/>',
    a4,
    a4 + '<w:headerReference w:type="even" r:id="newEven"/><w:titlePg/>',
  ]);
  const r = resolveWordSections(c);
  expect(r[1].headers.default).toEqual({
    relationshipId: 'odd',
    sourceSectionId: r[0].id,
    inherited: true,
  });
  expect(r[2].headers.default).toEqual(r[1].headers.default);
  expect(r[2].headers.even).toEqual({
    relationshipId: 'newEven',
    sourceSectionId: r[2].id,
    inherited: false,
  });
  expect(r[2].footers.first?.relationshipId).toBe('foot');
  expect(r[2].differentFirstPage).toBe(true);
  expect(r[0].headers.first).toBeNull();
});
it('does not read historical page geometry or historical first-page flags', () => {
  const c = content([
    a4 + '<w:sectPrChange><w:sectPr>' + letter + '<w:titlePg/></w:sectPr></w:sectPrChange>',
  ]);
  expect(resolveWordSections(c)[0]).toMatchObject({
    width: 16838,
    height: 11906,
    differentFirstPage: false,
  });
});
it('applies global orientation and margin overrides to each section without mutating source data', () => {
  const c = content([letter, a4]),
    source = JSON.stringify(c.docxStructure);
  c.orientation = 'portrait';
  c.margin = 'narrow';
  const result = resolveWordSections(c);
  expect(result.map((s) => [s.width, s.height])).toEqual([
    [12240, 15840],
    [11906, 16838],
  ]);
  expect(result[0].margins).toMatchObject({
    left: 720,
    right: 720,
    top: 720,
    bottom: 720,
    header: 400,
    footer: 500,
  });
  expect(JSON.stringify(c.docxStructure)).toBe(source);
});
it('paper overrides retain the existing whole-document command orientation semantics', () => {
  const c = content([letter, a4]);
  c.paper = 'letter';
  expect(resolveWordSections(c).map((s) => [s.width, s.height])).toEqual([
    [15840, 12240],
    [15840, 12240],
  ]);
});
it('keeps absent or invalid geometry unknown instead of claiming native defaults', () => {
  const c = content(['<w:pgSz w:w="oops" w:h="0"/><w:pgMar w:top="-500"/>']);
  c.orientation = 'portrait';
  expect(resolveWordSections(c)[0]).toMatchObject({
    width: null,
    height: null,
    margins: { top: -500, left: null },
  });
  expect(singleSectionPage(c)).toBeNull();
});
it('offers exact single-section rectangles only when required geometry is known and supported', () => {
  const c = content([letter]);
  c.paper = 'letter';
  c.orientation = 'portrait';
  expect(singleSectionPage(c)).toMatchObject({ width: 12240, height: 15840 });
  expect(singleSectionPage(content([letter, a4]))).toBeNull();
  // Without retained source metadata the model's global page controls define
  // the single authored section.
  expect(singleSectionPage({ ...c, docxStructure: undefined })).toMatchObject({
    width: 12240,
    height: 15840,
  });
  const gutter = content([letter.replace('w:gutter="0"', 'w:gutter="500"')]);
  gutter.paper = 'letter';
  gutter.orientation = 'portrait';
  expect(singleSectionPage(gutter)).toBeNull();
});
it('bounds untrusted section XML and rejects DTDs before parsing', () => {
  const c = content([a4]);
  c.docxStructure!.sections[0].propertiesXml = '<!DOCTYPE x [<!ENTITY y "z">]><x/>';
  expect(() => resolveWordSections(c)).toThrow('DTD');
});

it('explicit presets apply even when the old import approximation already has the same value', () => {
  const c = content([letter, a4]);
  c.pageOverrides = { paper: true, margin: true };
  expect(resolveWordSections(c).map((s) => [s.width, s.height, s.margins.right])).toEqual([
    [16838, 11906, 1440],
    [16838, 11906, 1440],
  ]);
});
