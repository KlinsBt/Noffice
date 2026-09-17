import type { Node } from '@tiptap/pm/model';
import { wordHyphenText } from './word-hyphen';
import type { WordSectionMap } from './word-section-ranges';
import type { WordSectionLayout } from './word-section-layout';
import type { WordStoryMeasurements, WordStoryPlacement } from './word-story-layout';
import type { WordNumbering } from './word-list-layout';
import { resolveWordLists, type WordListParagraph } from './word-list-resolve';
import { sectionParagraphs } from './word-section-breaks';

export interface SurfaceBlock {
  from: number;
  to: number;
  section: number;
  width: number;
  /** Physical leading margin, retained by screen CSS and explicit in print. */
  indentStart?: number;
  list?: WordListParagraph;
  keepNext?: boolean;
  pageBreakBefore?: boolean;
  flowBreaks?: boolean;
  pageBreakParagraph?: boolean;
}
export interface SurfaceInput {
  sections: WordSectionLayout[];
  blocks: SurfaceBlock[];
  width: number;
  stories?: WordStoryMeasurements;
  storyHtml?: Record<string, string>;
  storyText?: Record<string, string>;
}
export interface BlockMeasurement {
  height: number;
  before: number;
  after: number;
}
export interface WordSurfacePlan {
  width: number;
  height: number;
  pages: {
    sectionId: string;
    left: number;
    top: number;
    width: number;
    height: number;
    blank?: boolean;
    stories?: WordStoryPlacement[];
    bodyTop?: number;
    bodyBottom?: number;
  }[];
  blocks: (SurfaceBlock & { left: number; top: number })[];
  separators?: { page: number; left: number; top: number; height: number }[];
  tabBars?: { page: number; left: number; top: number; width: number; height: number }[];
}

/** Eligibility precedes measurement. Never infer one page per arbitrary section.
 * This path covers top-level paragraphs, known geometry, columns and
 * measured flow controls. Side stories remain separate dependencies; overflow
 * is checked using real line measurements after this structural eligibility gate.
 */
export function sectionSurfaceInput(
  doc: Node,
  map: WordSectionMap | undefined,
  sections: WordSectionLayout[],
  stories?: WordStoryMeasurements,
  measuredTabParagraphs?: ReadonlySet<number>,
  numbering?: WordNumbering,
): SurfaceInput | null {
  if (
    !map ||
    map.issues.length ||
    sections.length < 1 ||
    sections.length > 100 ||
    sections.length !== map.ranges.length ||
    map.size !== doc.content.size
  )
    return null;
  if (
    sections.some(
      (s, i) =>
        s.id !== map.ranges[i].id ||
        !s.simpleBody ||
        // The pinned Word build lays out imported nextColumn on a new page,
        // including a partially filled column. Keep its source type intact;
        // native boundary fixtures cover available/full/overflowing columns.
        (i > 0 &&
          !['nextPage', 'nextColumn', 'continuous', 'oddPage', 'evenPage'].includes(s.start)) ||
        !s.width ||
        !s.height ||
        s.margins.gutter !== 0 ||
        (!stories &&
          (Object.values(s.headers).some(Boolean) || Object.values(s.footers).some(Boolean))) ||
        [s.margins.top, s.margins.right, s.margins.bottom, s.margins.left].some(
          (n) => n === null || n < 0,
        ) ||
        s.margins.left! + s.margins.right! >= s.width ||
        s.margins.top! + s.margins.bottom! >= s.height,
    )
  )
    return null;
  const blocks: SurfaceBlock[] = [];
  const lists = resolveWordLists(doc, numbering);
  if (!lists) return null;
  let valid = doc.childCount > 0,
    section = 0;
  // Paragraphs must be top-level or inside a qualified flat source list. Table,
  // heading and nested-container paths stay outside this layout contract.
  doc.forEach(node => {
    if (node.type.name !== 'paragraph' && !['orderedList', 'bulletList'].includes(node.type.name)) valid = false;
  });
  sectionParagraphs(doc).forEach(({ node, from, topLevel }) => {
    const list = lists.get(from);
    if (!topLevel && !list) valid = false;
    while (section < sections.length - 1 && from >= map.ranges[section].to) section++;
    const range = map.ranges[section],
      layout = sections[section];
    const physicalIndent = (value: unknown) => {
      if (value == null || value === '') return 0;
      const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))(pt|px)?$/.exec(String(value));
      return match ? Number(match[1]) * (match[2] === 'px' ? 1 : 4 / 3) : NaN;
    };
    const startIndent = list ? list.definition.left / 15 : physicalIndent(node.attrs.indentStart),
      endIndent = physicalIndent(node.attrs.indentEnd),
      firstIndent = physicalIndent(node.attrs.firstLineIndent);
    const bodyWidth =
      (layout.columns
        ? Math.min(...layout.columns.widths)
        : layout.width! - layout.margins.left! - layout.margins.right!) / 15;
    if (
      node.type.name !== 'paragraph' ||
      from < range.from ||
      from + node.nodeSize > range.to ||
      /^[12]:/.test(node.attrs.sourceParagraph || '') ||
      node.attrs.direction === 'rtl' ||
      ![startIndent, endIndent, firstIndent].every(Number.isFinite) ||
      startIndent < 0 ||
      endIndent < 0 ||
      Math.max(startIndent, endIndent, Math.abs(firstIndent)) > 960 ||
      startIndent + endIndent >= bodyWidth ||
      startIndent + firstIndent < 0 ||
      startIndent + endIndent + firstIndent >= bodyWidth ||
      ((startIndent !== 0 || endIndent !== 0 || firstIndent !== 0) && !!layout.columns)
    )
      valid = false;
    // Until width-dependent reflow is available, unequal columns share a
    // measured minimum width and require left alignment. Line measurement
    // subsequently rejects soft wrapping instead of inventing a shared wrap.
    const unequal = layout.columns?.widths.some(
      (w) => Math.abs(w - layout.columns!.widths[0]) > 0.001,
    );
    if (unequal && node.attrs.textAlign && node.attrs.textAlign !== 'left') valid = false;
    let flowBreaks = false;
    node.forEach((inline) => {
      if (['wordPageBreak', 'wordColumnBreak'].includes(inline.type.name)) flowBreaks = true;
      else if (
        inline.type.name === 'wordTab' &&
        measuredTabParagraphs?.has(from) &&
        startIndent === 0 &&
        firstIndent === 0
      ) {
        /* Verified physical line. */
      } else if (inline.type.name === 'wordHyphen' && wordHyphenText(inline.attrs.kind) !== null) {
        /* Measured inline character. */
      } else if (!inline.isText && inline.type.name !== 'hardBreak') valid = false;
    });
    blocks.push({
      from,
      to: from + node.nodeSize,
      section,
      width: bodyWidth - startIndent - endIndent,
      ...(list ? { list } : {}),
      ...(startIndent ? { indentStart: startIndent } : {}),
      ...(node.attrs.keepNext ? { keepNext: true } : {}),
      ...(node.attrs.pageBreakBefore ? { pageBreakBefore: true } : {}),
      ...(flowBreaks ? { flowBreaks: true } : {}),
      ...(node.childCount === 1 &&
      node.firstChild?.type.name === 'wordPageBreak' &&
      from + node.nodeSize < doc.content.size
        ? { pageBreakParagraph: true }
        : {}),
    });
  });
  if (!valid || sections.some((_, i) => !blocks.some((b) => b.section === i))) return null;
  return {
    sections,
    blocks,
    width: Math.max(...sections.map((s) => s.width! / 15)),
    ...(stories ? { stories } : {}),
  };
}

/** Measured short-flow pagination. Coordinates are CSS pixels at 100% zoom.
 * Margin collapse is explicit; each section must fit inside its own body rectangle.
 * Oversized sections return null, never a clipped or invented one-page result.
 */
export function planSectionSurfaces(
  input: SurfaceInput,
  measurements: BlockMeasurement[],
  gap = 24,
): WordSurfacePlan | null {
  if (
    input.stories ||
    input.sections.some(
      (s, i) => s.columns || (i > 0 && ['continuous', 'oddPage', 'evenPage'].includes(s.start)),
    ) ||
    input.blocks.some((b) => b.flowBreaks) ||
    measurements.length !== input.blocks.length ||
    !Number.isFinite(gap) ||
    gap < 0 ||
    measurements.some(
      (m) =>
        !Number.isFinite(m.height) ||
        m.height <= 0 ||
        !Number.isFinite(m.before) ||
        m.before < 0 ||
        !Number.isFinite(m.after) ||
        m.after < 0,
    )
  )
    return null;
  const result: WordSurfacePlan = { width: input.width, height: 0, pages: [], blocks: [] };
  let index = 0;
  for (let section = 0; section < input.sections.length; section++) {
    const s = input.sections[section],
      top = result.height,
      width = s.width! / 15,
      height = s.height! / 15,
      left = (input.width - width) / 2;
    result.pages.push({ sectionId: s.id, left, top, width, height });
    let y = s.margins.top! / 15,
      previousAfter = 0;
    while (index < input.blocks.length && input.blocks[index].section === section) {
      const b = input.blocks[index],
        m = measurements[index];
      if (b.pageBreakBefore && index > 0 && input.blocks[index - 1].section === section)
        return null;
      y += Math.max(previousAfter, m.before);
      if (y + m.height > height - s.margins.bottom! / 15 + 0.01) return null;
      // Absolute CSS offsets include the element's own top margin.
      result.blocks.push({ ...b, left: left + s.margins.left! / 15, top: top + y - m.before });
      y += m.height;
      previousAfter = m.after;
      index++;
    }
    result.height += height + (section < input.sections.length - 1 ? gap : 0);
  }
  return result;
}

/** A continuous start shares the current physical page only when its page/body
 * geometry is unchanged. Different native page dimensions start a new page. */
export function continuesSectionPage(input: SurfaceInput, index: number): boolean {
  if (index === 0 || input.sections[index].start !== 'continuous') return false;
  const a = input.sections[index - 1],
    b = input.sections[index];
  return (
    a.width === b.width &&
    a.height === b.height &&
    a.orientation === b.orientation &&
    ((!a.columns && !b.columns && a.margins.gutter === 0 && b.margins.gutter === 0) ||
      (['top', 'right', 'bottom', 'left', 'gutter'] as const).every(
        (side) => a.margins[side] === b.margins[side],
      ))
  );
}

export function sectionPrintPage(input: SurfaceInput, index: number): number {
  while (continuesSectionPage(input, index)) index--;
  return index;
}

/** Same resolved section rectangles drive named print pages; no layout CSS enters saved HTML. */
export function sectionPrintRules(input: SurfaceInput): string {
  return input.sections
    .map(
      (s, i) =>
        `@page nofficeSection${i} { size: ${s.width! / 20}pt ${s.height! / 20}pt; margin: ${['top', 'right', 'bottom', 'left'].map((side) => `${s.margins[side as 'top']! / 20}pt`).join(' ')}; }\n` +
        // A fragment tree already owns physical margins and page allocation.
        // Giving it the full paper avoids print shrinkage from hanging spaces
        // crossing a CSS content margin; clipping occurs only at the paper edge.
        `@page nofficeFragment${i} { size: ${s.width! / 20}pt ${s.height! / 20}pt; margin: 0; }`,
    )
    .join('\n');
}
