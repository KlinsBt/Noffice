import { it, expect, vi } from 'vitest';
import { Document, Packer, Paragraph } from 'docx';
import { readDocxStructure, liveSectionProperties, docxStructureSchema } from './docx-sections';
import { wordXml, WORD_NS } from './docx-import';
import { contentSchema, newFile } from './model';
import { hydrateWordStructure } from './word-structure';
import { contentFingerprint } from './office-preservation';
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
const xml = (body: string) =>
  wordXml(
    `<w:document xmlns:w="${WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><w:body>${body}</w:body></w:document>`,
  );
const fixture = () =>
  xml(
    '<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:headerReference w:type="even" r:id="rEven"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:footnoteReference w:id="2"/><w:drawing><wp:anchor><wp:docPr id="7"/><a:blip r:embed="rImage"/></wp:anchor></w:drawing></w:r></w:p><w:sectPr><w:footerReference w:type="default" r:id="rFooter"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>',
  );
it('maps every main-story paragraph to stable source sections, including table cells', () => {
  const a = readDocxStructure(fixture()),
    b = readDocxStructure(fixture());
  expect(a).toEqual(b);
  expect(a.sections.map((s) => s.paragraphs)).toEqual([
    ['0:0', '0:1'],
    ['0:2', '0:3'],
  ]);
  expect(a.sections.map((s) => s.endingParagraph)).toEqual(['0:1', null]);
  expect(a.sections[0].headers).toEqual([{ type: 'even', relationshipId: 'rEven' }]);
  expect(a.sections[1].footers).toEqual([{ type: 'default', relationshipId: 'rFooter' }]);
  expect(a.sections[1].notes).toEqual([{ paragraph: '0:3', kind: 'footnote', id: '2' }]);
  expect(a.sections[1].drawings).toEqual([
    { paragraph: '0:3', id: '7', relationshipIds: ['rImage'] },
  ]);
});
it('ignores historical, table and textbox section markers', () => {
  const d = xml(
    '<w:p><w:pPr><w:pPrChange><w:pPr><w:sectPr/></w:pPr></w:pPrChange></w:pPr><w:r><w:txbxContent><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:txbxContent></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:sectPrChange><w:sectPr><w:pgSz w:w="16838" w:h="11906"/></w:sectPr></w:sectPrChange></w:sectPr>',
  );
  expect(liveSectionProperties(d.documentElement.firstElementChild!)).toHaveLength(1);
  expect(readDocxStructure(d).sections).toHaveLength(1);
  expect(readDocxStructure(d).sections[0].paragraphs).toHaveLength(3);
});
it('keeps nested drawing text in the section that ends at its containing paragraph', () => {
  const s = readDocxStructure(
    xml(
      '<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:txbxContent><w:p/></w:txbxContent></w:r></w:p><w:p/><w:sectPr/>',
    ),
  ).sections;
  expect(s[0].paragraphs).toEqual(['0:0', '0:1']);
  expect(s[1].paragraphs).toEqual(['0:2']);
});
it('represents implicit final settings without inventing source XML', () => {
  expect(readDocxStructure(xml('<w:p/>')).sections[0]).toMatchObject({
    paragraphs: ['0:0'],
    propertiesXml: null,
    endingParagraph: null,
  });
});
it('retains content controls as section containers without counting deleted breaks', () => {
  const d = xml(
    '<w:sdt><w:sdtContent><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:sdtContent></w:sdt><w:del><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:del><w:sectPr/>',
  );
  expect(readDocxStructure(d).sections).toHaveLength(2);
});
it('round trips additive metadata through model validation and rejects future versions', () => {
  const file = newFile('word');
  if (file.content.kind !== 'word') throw Error();
  const content = { ...file.content, docxStructure: readDocxStructure(fixture()) };
  expect(contentSchema.parse(JSON.parse(JSON.stringify(content)))).toEqual(content);
  expect(() => docxStructureSchema.parse({ ...content.docxStructure, version: 2 })).toThrow();
  expect(contentSchema.parse(file.content)).toEqual(file.content);
});
it('hydrates old snapshots atomically without replacing HTML, revision or original bytes', async () => {
  const data = await Packer.toArrayBuffer(
      new Document({
        // This case isolates section hydration with an unaffected explicit
        // cascade. Ambiguous old application defaults have separate guards.
        styles: {
          default: {
            document: {
              run: { font: 'Arial', size: 22 },
              paragraph: { spacing: { line: 240, after: 0 } },
            },
          },
        },
        sections: [{ children: [new Paragraph('Original')] }],
      }),
    ),
    file = newFile('word');
  file.original = { name: 'source.docx', data };
  if (file.content.kind !== 'word') throw Error();
  file.content.html = '<p>Unsaved edit</p>';
  const beforeFingerprint = await contentFingerprint(file.content);
  const upgraded = await hydrateWordStructure(file);
  expect(upgraded.content).toMatchObject({
    html: '<p>Unsaved edit</p>',
    docxStructure: { version: 1 },
  });
  expect(upgraded.original).toBe(file.original);
  expect(upgraded.revision).toBe(file.revision);
  expect(file.content.docxStructure).toBeUndefined();
  expect(await contentFingerprint(upgraded.content)).toBe(beforeFingerprint);
  expect(await hydrateWordStructure(upgraded)).toBe(upgraded);
});
it('rejects malformed legacy originals without mutating the saved edit', async () => {
  const file = newFile('word');
  file.original = { name: 'x.docx', data: new ArrayBuffer(2) };
  await expect(hydrateWordStructure(file)).rejects.toThrow();
  expect(file.content).not.toHaveProperty('docxStructure');
});
