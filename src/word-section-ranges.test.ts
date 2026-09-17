import { describe, expect, it } from 'vitest';
import { Editor, getSchema } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import type { DocxStructure } from './docx-sections';
import { mapWordSectionRanges, selectedWordSections } from './word-section-ranges';

const schema = getSchema(wordExtensions());
const p = (id: string | null, text = '') =>
  schema.nodes.paragraph.create({ sourceParagraph: id }, text ? schema.text(text) : undefined);
const doc = (...nodes: ReturnType<typeof p>[]) => schema.nodes.doc.create(null, nodes);
const structure: DocxStructure = {
  version: 1,
  sourcePath: 'word/document.xml',
  sections: [
    {
      id: 'word/document.xml#section:0',
      endingParagraph: '0:1',
      paragraphs: ['0:0', '0:1'],
      propertiesXml: null,
      headers: [],
      footers: [],
      drawings: [],
      notes: [],
    },
    {
      id: 'word/document.xml#section:1',
      endingParagraph: null,
      paragraphs: ['0:2', '0:3'],
      propertiesXml: null,
      headers: [],
      footers: [],
      drawings: [],
      notes: [],
    },
  ],
};
const ids = structure.sections.map((s) => s.id);

describe('editor source-section ranges', () => {
  it('includes the empty ending paragraph in its section and uses half-open positions', () => {
    const mapped = mapWordSectionRanges(doc(p('0:0', 'One'), p('0:1'), p('0:2', 'Two')), structure);
    expect(mapped.issues).toEqual([]);
    expect(mapped.ranges).toEqual([
      { id: ids[0], from: 0, to: 7 },
      { id: ids[1], from: 7, to: 12 },
    ]);
    expect(selectedWordSections(mapped, 6, 6)).toEqual([ids[0]]);
    expect(selectedWordSections(mapped, 7, 7)).toEqual([ids[1]]);
    expect(selectedWordSections(mapped, 1, 7)).toEqual([ids[0]]);
    expect(selectedWordSections(mapped, 6, 8)).toEqual(ids);
    expect(selectedWordSections(mapped, 8, 6)).toEqual(ids);
    expect(selectedWordSections(mapped, 12, 12)).toEqual([ids[1]]);
  });
  it('uses document offsets inside lists and tables without inventing extra sections', () => {
    const list = schema.nodes.bulletList.create(
      null,
      schema.nodes.listItem.create(null, p('0:0', 'Item')),
    );
    const table = schema.nodes.table.create(
      null,
      schema.nodes.tableRow.create(null, schema.nodes.tableCell.create(null, p('0:2', 'Cell'))),
    );
    const document = doc(list, p('0:1'), table, p('0:3'));
    const mapped = mapWordSectionRanges(document, structure);
    expect(mapped.issues).toEqual([]);
    expect(mapped.paragraphs.map((p) => [p.from, p.to, p.sectionId])).toEqual([
      [2, 8, ids[0]],
      [10, 12, ids[0]],
      [15, 21, ids[1]],
      [24, 26, ids[1]],
    ]);
    expect(selectedWordSections(mapped, 16, 17)).toEqual([ids[1]]);
  });
  it('assigns new paragraphs by surviving boundaries, including before the first source paragraph', () => {
    const mapped = mapWordSectionRanges(
      doc(p(null, 'New'), p('0:0'), p(null), p('0:1'), p(null, 'After'), p('0:2')),
      structure,
    );
    expect(mapped.issues).toEqual([]);
    expect(mapped.paragraphs.map((p) => p.sectionId)).toEqual([
      ids[0],
      ids[0],
      ids[0],
      ids[0],
      ids[1],
      ids[1],
    ]);
  });
  it('keeps imported note stories out of the main-body section selection', () => {
    const mapped = mapWordSectionRanges(
      doc(p('0:0'), p('0:1'), p('0:2'), p('1:0', 'Note'), p('2:0', 'Endnote')),
      structure,
    );
    expect(mapped.issues).toEqual([]);
    expect(mapped.paragraphs.slice(-2).map((p) => p.sectionId)).toEqual([null, null]);
    expect(selectedWordSections(mapped, 7, 8)).toEqual([]);
    expect(selectedWordSections(mapped, 5, 9)).toEqual([ids[1]]);
  });
  it('does not infer missing, repeated or reordered source boundaries', () => {
    for (const [document, code] of [
      [doc(p('0:0'), p('0:2')), 'missing-boundary'],
      [doc(p('0:0'), p('0:1'), p('0:1'), p('0:2')), 'duplicate-boundary'],
      [doc(p('0:2'), p('0:1'), p('0:0')), 'source-order'],
      [doc(p('0:0'), p('0:1'), p('0:99')), 'unknown-source'],
    ] as const) {
      const mapped = mapWordSectionRanges(document, structure);
      expect(mapped.issues.map((i) => i.code)).toContain(code);
      expect(mapped.ranges).toEqual([]);
      expect(selectedWordSections(mapped, 1, 1)).toEqual([]);
    }
  });
  it('rejects malformed source identities without modifying source metadata', () => {
    const broken = structuredClone(structure);
    broken.sections[1].paragraphs.push('0:1');
    const before = JSON.stringify(broken);
    expect(mapWordSectionRanges(doc(p('0:1')), broken).issues.map((i) => i.code)).toContain(
      'invalid-source',
    );
    expect(JSON.stringify(broken)).toBe(before);
    expect(mapWordSectionRanges(doc(p(null)), undefined).ranges).toEqual([]);
  });
  it('allows ordinary split/join and recomputes from immutable editor history', () => {
    const editor = new Editor({
      extensions: wordExtensions(),
      content:
        '<p data-source-paragraph="0:0">AlphaBeta</p><p data-source-paragraph="0:1"></p><p data-source-paragraph="0:2">After</p>',
    });
    try {
      const initial = mapWordSectionRanges(editor.state.doc, structure);
      editor.commands.setTextSelection(6);
      editor.commands.splitBlock();
      const split = mapWordSectionRanges(editor.state.doc, structure);
      expect(split.issues).toEqual([]);
      expect(split.paragraphs.map((p) => p.sectionId)).toEqual([ids[0], ids[0], ids[0], ids[1]]);
      expect(split.ranges[0].to).toBe(initial.ranges[0].to + 2);
      editor.commands.undo();
      expect(mapWordSectionRanges(editor.state.doc, structure)).toEqual(initial);
      editor.commands.redo();
      expect(mapWordSectionRanges(editor.state.doc, structure)).toEqual(split);
      editor.commands.joinBackward();
      expect(mapWordSectionRanges(editor.state.doc, structure)).toEqual(initial);
    } finally {
      editor.destroy();
    }
  });
  it('reports boundary split/deletion as unresolved and restores mapping on undo', () => {
    const editor = new Editor({
      extensions: wordExtensions(),
      content:
        '<p data-source-paragraph="0:0">First</p><p data-source-paragraph="0:1">Break</p><p data-source-paragraph="0:2">Last</p>',
    });
    try {
      const initial = mapWordSectionRanges(editor.state.doc, structure);
      editor.commands.setTextSelection(10);
      editor.commands.splitBlock();
      expect(mapWordSectionRanges(editor.state.doc, structure).issues.map((i) => i.code)).toContain(
        'duplicate-boundary',
      );
      editor.commands.undo();
      expect(mapWordSectionRanges(editor.state.doc, structure)).toEqual(initial);
      editor.commands.deleteRange({ from: 7, to: 14 });
      expect(mapWordSectionRanges(editor.state.doc, structure).issues.map((i) => i.code)).toContain(
        'missing-boundary',
      );
      editor.commands.undo();
      expect(mapWordSectionRanges(editor.state.doc, structure)).toEqual(initial);
    } finally {
      editor.destroy();
    }
  });
  it('rejects invalid selection coordinates and gives shared boundary affinity explicitly', () => {
    const mapped = mapWordSectionRanges(doc(p('0:0'), p('0:1'), p('0:2')), structure);
    expect(selectedWordSections(mapped, 4, 4, -1)).toEqual([ids[0]]);
    for (const position of [-1, 1.5, NaN, Infinity, 7])
      expect(selectedWordSections(mapped, position, position)).toEqual([]);
  });
});
