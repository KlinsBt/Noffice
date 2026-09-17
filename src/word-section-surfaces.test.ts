import { expect, it } from 'vitest';
import { getSchema } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { readDocxStructure } from './docx-sections';
import { wordXml, WORD_NS } from './docx-import';
import { mapWordSectionRanges } from './word-section-ranges';
import { resolveWordSections } from './word-section-layout';
import {
  sectionSurfaceInput,
  planSectionSurfaces,
  sectionPrintRules,
} from './word-section-surfaces';

function fixture(extra = '') {
  const geometry = (w: number, h: number) =>
    `<w:pgSz w:w="${w}" w:h="${h}" ${w > h ? 'w:orient="landscape"' : ''}/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/>`;
  const source = readDocxStructure(
    wordXml(
      `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:pPr><w:sectPr>${geometry(12240, 15840)}</w:sectPr></w:pPr></w:p><w:p/><w:sectPr>${geometry(16838, 11906)}${extra}</w:sectPr></w:body></w:document>`,
    ),
  );
  const schema = getSchema(wordExtensions());
  const doc = schema.nodeFromJSON({
    type: 'doc',
    content: ['Alpha', 'Beta'].map((text, i) => ({
      type: 'paragraph',
      attrs: { sourceParagraph: `0:${i}` },
      content: [{ type: 'text', text }],
    })),
  });
  const sections = resolveWordSections({
    kind: 'word',
    html: '',
    paper: 'a4',
    margin: 'normal',
    orientation: 'landscape',
    docxStructure: source,
  });
  const map = mapWordSectionRanges(doc, source);
  return { doc, sections, map };
}
const metrics = [
  { height: 20, before: 0, after: 12 },
  { height: 30, before: 0, after: 12 },
];
it('reserves both physical indents while first/hanging indentation affects only the first line', () => {
  const f = fixture(), json = f.doc.toJSON();
  for (const first of ['5pt', '-12pt']) {
    Object.assign(json.content[0].attrs, { indentStart: '12pt', indentEnd: '6pt', firstLineIndent: first });
    const input = sectionSurfaceInput(f.doc.type.schema.nodeFromJSON(json), f.map, f.sections)!;
    expect(input.blocks[0]).toMatchObject({ width: 600, indentStart: 16 });
    expect(planSectionSurfaces(input, metrics)!.blocks[0].width).toBe(600);
  }
  for (const attrs of [{ firstLineIndent: '-13pt' }, { firstLineIndent: '450pt' },
    { indentStart: '462pt' }, { indentStart: 'NaN' }, { indentEnd: 'Infinity' }]) {
    Object.assign(json.content[0].attrs, { indentStart: '12pt', indentEnd: '6pt', firstLineIndent: '5pt' }, attrs);
    expect(sectionSurfaceInput(f.doc.type.schema.nodeFromJSON(json), f.map, f.sections)).toBeNull();
  }
});
it('uses the reduced right-indented body width for measurement and actual page fragments', () => {
  const f = fixture(), json = f.doc.toJSON();
  for (const value of ['36pt', '48px']) {
    json.content[0].attrs.indentEnd = value;
    const input = sectionSurfaceInput(f.doc.type.schema.nodeFromJSON(json), f.map, f.sections)!;
    expect(input.blocks[0].width).toBe(576);
    const plan = planSectionSurfaces(input, metrics)!;
    expect(plan.blocks[0].width).toBe(576);
    expect(plan.blocks[0].left).toBeCloseTo((16838 / 15 - 816) / 2 + 96);
  }
  for (const value of ['-1pt', '468pt', '721pt', '12%', 'invalid', 'Infinity']) {
    json.content[0].attrs.indentEnd = value;
    expect(sectionSurfaceInput(f.doc.type.schema.nodeFromJSON(json), f.map, f.sections)).toBeNull();
  }
});
it('requires measured tab provenance and permits explicit zero indents without accepting nonzero or invalid values', () => {
  const f = fixture();
  const json = f.doc.toJSON();
  json.content[0].content = [{ type: 'text', text: 'A' }, { type: 'wordTab' }, { type: 'text', text: 'pha' }];
  const schema = f.doc.type.schema;
  expect(sectionSurfaceInput(schema.nodeFromJSON(json), f.map, f.sections)).toBeNull();
  for (const value of ['0pt', '0px', '-0.00pt']) {
    json.content[0].attrs.indentStart = value;
    json.content[0].attrs.indentEnd = value;
    json.content[0].attrs.firstLineIndent = value;
    expect(sectionSurfaceInput(schema.nodeFromJSON(json), f.map, f.sections, undefined, new Set([0]))).not.toBeNull();
  }
  for (const value of ['1pt', '-1pt', 'invalid', '0unknown']) {
    json.content[0].attrs.indentStart = value;
    expect(sectionSurfaceInput(schema.nodeFromJSON(json), f.map, f.sections, undefined, new Set([0]))).toBeNull();
  }
});
it('places differently sized section bodies inside centered page rectangles without mutating source positions', () => {
  const f = fixture(),
    original = JSON.stringify(f.map);
  const input = sectionSurfaceInput(f.doc, f.map, f.sections)!;
  const plan = planSectionSurfaces(input, metrics)!;
  expect(plan.pages.map((p) => [p.width, p.height, p.top])).toEqual([
    [816, 1056, 0],
    [16838 / 15, 11906 / 15, 1080],
  ]);
  expect(plan.blocks[0].left).toBeCloseTo((16838 / 15 - 816) / 2 + 96);
  expect(plan.blocks[1]).toMatchObject({ left: 96, top: 1176 });
  expect(plan.blocks.map((b) => b.from)).toEqual([0, 7]);
  expect(JSON.stringify(f.map)).toBe(original);
  expect(sectionPrintRules(input)).toContain(
    '@page nofficeSection1 { size: 841.9pt 595.3pt; margin: 72pt 72pt 72pt 72pt; }',
  );
});
it.each(['oddPage', 'evenPage'])(
  'routes %s sections through physical fragment pagination',
  (start) => {
    const f = fixture(`<w:type w:val="${start}"/>`);
    const input = sectionSurfaceInput(f.doc, f.map, f.sections);
    expect(input).not.toBeNull();
    expect(planSectionSurfaces(input!, metrics)).toBeNull();
  },
);
it('accepts equal columns but declines headers, unknown margins and unresolved boundaries', () => {
  const f = fixture('<w:cols w:num="2"/>');
  expect(sectionSurfaceInput(f.doc, f.map, f.sections)).not.toBeNull();
  expect(f.sections[1].columns?.widths).toHaveLength(2);
  for (const change of [
    (x: typeof f) => {
      x.sections[1].margins.top = null;
    },
    (x: typeof f) => {
      x.sections[1].headers.default = {
        relationshipId: 'rId1',
        sourceSectionId: 'section',
        inherited: true,
      };
    },
    (x: typeof f) => {
      x.map.issues.push({ code: 'missing-boundary', sourceId: '0:0' });
    },
  ]) {
    const x = fixture();
    change(x);
    expect(sectionSurfaceInput(x.doc, x.map, x.sections)).toBeNull();
  }
});
it('declines tables, notes and unsupported paragraph indents', () => {
  for (const attrs of [{ indentStart: '-12pt' }, { sourceParagraph: '1:0' }]) {
    const f = fixture(),
      json = f.doc.toJSON();
    Object.assign(json.content[0].attrs, attrs);
    expect(sectionSurfaceInput(f.doc.type.schema.nodeFromJSON(json), f.map, f.sections)).toBeNull();
  }
  const f = fixture(),
    json = f.doc.toJSON();
  json.content[0].type = 'heading';
  expect(sectionSurfaceInput(f.doc.type.schema.nodeFromJSON(json), f.map, f.sections)).toBeNull();
});
it('rejects overflowing paragraphs instead of clipping the page, and recovers after shortening', () => {
  const f = fixture(),
    input = sectionSurfaceInput(f.doc, f.map, f.sections)!;
  expect(planSectionSurfaces(input, [{ ...metrics[0], height: 865 }, metrics[1]])).toBeNull();
  expect(planSectionSurfaces(input, [{ ...metrics[0], height: 864 }, metrics[1]])).not.toBeNull();
  expect(planSectionSurfaces(input, metrics)).not.toBeNull();
});
it('rejects incomplete/nonfinite/negative measurements', () => {
  const f = fixture(),
    input = sectionSurfaceInput(f.doc, f.map, f.sections)!;
  expect(planSectionSurfaces(input, [])).toBeNull();
  for (const m of [{ height: NaN }, { height: 0 }, { before: -1 }, { after: Infinity }])
    expect(planSectionSurfaces(input, [{ ...metrics[0], ...m }, metrics[1]])).toBeNull();
});
