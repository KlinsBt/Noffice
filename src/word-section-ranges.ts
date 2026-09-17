import type { Node } from '@tiptap/pm/model';
import type { DocxStructure } from './docx-sections';

export interface WordSectionRange {
  id: string;
  from: number;
  to: number;
}
export interface WordParagraphRange {
  from: number;
  to: number;
  sourceId: string | null;
  sectionId: string | null;
}
export interface WordSectionMap {
  size: number;
  ranges: WordSectionRange[];
  paragraphs: WordParagraphRange[];
  issues: {
    code:
      | 'invalid-source'
      | 'missing-boundary'
      | 'duplicate-boundary'
      | 'source-order'
      | 'unknown-source';
    sourceId: string | null;
  }[];
}

/** Derived positions, never persisted identities. Invalid boundaries fail closed for layout consumers.
 * Ordinary repeated IDs are split fragments (the retained writer keeps the first identity).
 * Repeated section-ending IDs are ambiguous until section-break transactions exist.
 */
export function mapWordSectionRanges(doc: Node, source: DocxStructure | undefined): WordSectionMap {
  const result: WordSectionMap = { size: doc.content.size, ranges: [], paragraphs: [], issues: [] };
  if (!source) return result;
  const issue = (code: WordSectionMap['issues'][number]['code'], sourceId: string | null) => {
    // Diagnostics are bounded independently of document size.
    if (result.issues.length < 100) result.issues.push({ code, sourceId });
  };
  const order = new Map<string, number>();
  const sectionIds = new Set<string>();
  for (let i = 0; i < source.sections.length; i++) {
    const section = source.sections[i];
    if (sectionIds.has(section.id)) issue('invalid-source', null);
    sectionIds.add(section.id);
    for (const id of section.paragraphs) {
      if (order.has(id)) issue('invalid-source', id);
      order.set(id, order.size);
    }
    if (
      i < source.sections.length - 1 &&
      (!section.endingParagraph || !section.paragraphs.includes(section.endingParagraph))
    )
      issue('invalid-source', section.endingParagraph);
    if (i === source.sections.length - 1 && section.endingParagraph !== null)
      issue('invalid-source', section.endingParagraph);
  }
  if (!source.sections.length || result.issues.length) return result;
  const occurrences = new Map<string, WordParagraphRange[]>();
  let previousOrder = -1;
  doc.descendants((node, pos) => {
    if (node.type.name !== 'paragraph' && node.type.name !== 'heading') return;
    const id = typeof node.attrs.sourceParagraph === 'string' ? node.attrs.sourceParagraph : null;
    const paragraph: WordParagraphRange = {
      from: pos,
      to: pos + node.nodeSize,
      sourceId: id,
      sectionId: null,
    };
    result.paragraphs.push(paragraph);
    if (id && !/^[12]:\d+$/.test(id)) {
      const found = occurrences.get(id);
      if (found) found.push(paragraph);
      else {
        occurrences.set(id, [paragraph]);
        const index = order.get(id);
        if (index === undefined) issue('unknown-source', id);
        else {
          if (index < previousOrder) issue('source-order', id);
          previousOrder = index;
        }
      }
    }
    return false;
  });
  let start = 0;
  for (const section of source.sections) {
    let end = doc.content.size;
    if (section.endingParagraph) {
      const endings = occurrences.get(section.endingParagraph);
      if (!endings?.length) {
        issue('missing-boundary', section.endingParagraph);
        continue;
      }
      if (endings.length !== 1) {
        issue('duplicate-boundary', section.endingParagraph);
        continue;
      }
      end = endings[0].to;
    }
    if (end < start) issue('source-order', section.endingParagraph);
    result.ranges.push({ id: section.id, from: start, to: end });
    start = end;
  }
  if (result.issues.length) {
    result.ranges = [];
    return result;
  }
  let sectionIndex = 0;
  for (const paragraph of result.paragraphs) {
    while (
      sectionIndex < result.ranges.length - 1 &&
      paragraph.from >= result.ranges[sectionIndex].to
    )
      sectionIndex++;
    if (!paragraph.sourceId || !/^[12]:\d+$/.test(paragraph.sourceId))
      paragraph.sectionId = result.ranges[sectionIndex].id;
  }
  return result;
}

/** Half-open selections. At a shared block boundary a caret belongs to the following section
 * by default; callers with backward association can request the preceding section.
 * Note paragraphs are a separate story, even when Mammoth appends them to the editor DOM.
 */
export function selectedWordSections(
  map: WordSectionMap,
  anchor: number,
  head: number,
  association: -1 | 1 = 1,
): string[] {
  if (
    ![anchor, head].every((p) => Number.isInteger(p) && p >= 0 && p <= map.size) ||
    map.issues.length
  )
    return [];
  const from = Math.min(anchor, head),
    to = Math.max(anchor, head);
  const notes = map.paragraphs.filter((p) => p.sectionId === null);
  if (from === to) {
    if (notes.some((p) => from >= p.from && from < p.to)) return [];
    const range = map.ranges.find((r) =>
      association === -1 ? from > r.from && from <= r.to : from >= r.from && from < r.to,
    );
    return range
      ? [range.id]
      : from === map.size
        ? map.ranges.slice(-1).map((r) => r.id)
        : from === 0
          ? map.ranges.slice(0, 1).map((r) => r.id)
          : [];
  }
  return map.ranges
    .filter((r) => {
      const left = Math.max(from, r.from),
        right = Math.min(to, r.to);
      if (left >= right) return false;
      let bodyLength = right - left;
      for (const note of notes)
        bodyLength -= Math.max(0, Math.min(right, note.to) - Math.max(left, note.from));
      return bodyLength > 0;
    })
    .map((r) => r.id);
}
