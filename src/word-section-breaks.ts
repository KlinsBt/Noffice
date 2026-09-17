import { z } from 'zod';
import type { Node } from '@tiptap/pm/model';
import { Mapping, ReplaceAroundStep, StepMap } from '@tiptap/pm/transform';
import type { Transaction } from '@tiptap/pm/state';
import type { DocxStructure } from './docx-sections';
import { mapWordSectionRanges, type WordSectionMap } from './word-section-ranges';
import { wordSectionGeometrySchema } from './word-section-geometry';
import { inFlatWordList } from './word-flat-list';
import {
  retainedWordSectionId,
  insertedWordSectionId,
  wordSectionIdSchema,
  wordSectionStartSchema,
} from './word-section-identity';

const identity = retainedWordSectionId;
const retainedState = z.object({
  version: z.literal(1),
  finalSectionId: identity,
  breaks: z
    .array(z.object({ sectionId: identity, paragraph: z.number().int().min(0).max(1000000) }))
    .max(9999),
});
const authoredState = z.object({
  version: z.literal(2),
  finalSectionId: wordSectionIdSchema,
  breaks: z
    .array(
      z.object({ sectionId: wordSectionIdSchema, paragraph: z.number().int().min(0).max(1000000) }),
    )
    .max(9999),
  inserted: z
    .array(
      z.object({
        id: insertedWordSectionId,
        sourceSectionId: z.union([retainedWordSectionId, z.literal('authored-body')]),
      }),
    )
    .max(9999),
  starts: z
    .array(z.object({ sectionId: wordSectionIdSchema, start: wordSectionStartSchema }))
    .max(10000),
  layouts: z.array(wordSectionGeometrySchema).max(10000).optional(),
});
export const wordSectionStateSchema = z.discriminatedUnion('version', [
  retainedState,
  authoredState,
]);
export type WordSectionState = z.infer<typeof wordSectionStateSchema>;
export type AuthoredWordSectionState = z.infer<typeof authoredState>;
export const liveWordSectionIds = (state: WordSectionState) => [
  ...state.breaks.map((b) => b.sectionId),
  state.finalSectionId,
];
export const wordSectionSourceId = (state: WordSectionState, id: string) =>
  state.version === 2 ? state.inserted.find((s) => s.id === id)?.sourceSectionId || id : id;

/** New identities refer only to immutable source sections (or the fresh body).
 * Never infer source ownership from an inserted section's current position. */
export function validateWordSectionIdentities(
  state: WordSectionState,
  source: DocxStructure | undefined,
) {
  wordSectionStateSchema.parse(state);
  if (!source && state.version === 1)
    throw Error('Retained sections require their original source.');
  const originals = source?.sections.map((s) => s.id) || ['authored-body'];
  if (state.finalSectionId !== originals.at(-1)) throw Error('Invalid final section identity.');
  const known = new Set(originals);
  if (state.version === 2) {
    for (const section of state.inserted) {
      if (known.has(section.id) || !originals.includes(section.sourceSectionId))
        throw Error('Invalid authored section provenance.');
      known.add(section.id);
    }
  }
  const live = liveWordSectionIds(state),
    used = new Set<string>();
  let prior = -1;
  for (const id of live) {
    if (!known.has(id) || used.has(id)) throw Error('Invalid live section identity.');
    used.add(id);
    const index = originals.indexOf(id);
    if (index >= 0) {
      if (index <= prior) throw Error('Invalid live section order.');
      prior = index;
    }
  }
  if (state.version === 2) {
    if (state.inserted.some((s) => !used.has(s.id)))
      throw Error('Unused authored section identity.');
    const overrides = new Set<string>();
    for (const s of state.starts) {
      if (!used.has(s.sectionId) || overrides.has(s.sectionId))
        throw Error('Invalid section start override.');
      overrides.add(s.sectionId);
    }
    const layouts = new Set<string>();
    for (const layout of state.layouts || []) {
      if (!used.has(layout.sectionId) || layouts.has(layout.sectionId))
        throw Error('Invalid section geometry override.');
      layouts.add(layout.sectionId);
    }
  }
  return live;
}
export function sectionParagraphs(doc: Node) {
  const paragraphs: { from: number; to: number; node: Node; topLevel: boolean; sectionBoundary: boolean }[] = [];
  doc.descendants((node, from, parent) => {
    if (!['paragraph', 'heading'].includes(node.type.name)) return;
    paragraphs.push({ from, to: from + node.nodeSize, node, topLevel: parent === doc,
      sectionBoundary: parent === doc || inFlatWordList(doc.resolve(from + 1)) });
    return false;
  });
  return paragraphs;
}
/** Paragraph ordinals refer to the normalized editable document, never to the retained source.
 * ProseMirror transactions transform boundaries; ordinals avoid persisting unstable text offsets.
 */
export function initializeSectionState(
  doc: Node,
  source: DocxStructure | undefined,
): WordSectionState | null {
  if (!source) return null;
  const mapped = mapWordSectionRanges(doc, source);
  if (!mapped.ranges.length || mapped.issues.length) return null;
  const paragraphs = sectionParagraphs(doc);
  const breaks = mapped.ranges.slice(0, -1).map((range) => ({
    sectionId: range.id,
    paragraph: paragraphs.findIndex((p) => p.to === range.to),
  }));
  if (breaks.some((b) => !paragraphs[b.paragraph]?.sectionBoundary)) return null;
  return { version: 1, finalSectionId: source.sections.at(-1)!.id, breaks };
}
export function validateSectionState(
  state: WordSectionState,
  source: DocxStructure | undefined,
  doc: Node,
) {
  validateWordSectionIdentities(state, source);
  const paragraphs = sectionParagraphs(doc);
  let previousParagraph = -1;
  for (const boundary of state.breaks) {
    if (boundary.paragraph <= previousParagraph || !paragraphs[boundary.paragraph]?.sectionBoundary)
      throw Error('Invalid live section boundary.');
    previousParagraph = boundary.paragraph;
  }
  return paragraphs;
}
export function mapLiveSections(
  doc: Node,
  source: DocxStructure | undefined,
  state: WordSectionState,
): WordSectionMap {
  const paragraphs = validateSectionState(state, source, doc);
  let from = 0;
  const ranges = state.breaks.map((b) => {
    const to = paragraphs[b.paragraph].to;
    const range = { id: b.sectionId, from, to };
    from = to;
    return range;
  });
  ranges.push({ id: state.finalSectionId, from, to: doc.content.size });
  let index = 0;
  return {
    size: doc.content.size,
    ranges,
    issues: [],
    paragraphs: paragraphs.map((p) => {
      while (index < ranges.length - 1 && p.from >= ranges[index].to) index++;
      const sourceId = p.node.attrs.sourceParagraph as string | null;
      return {
        from: p.from,
        to: p.to,
        sourceId,
        sectionId: sourceId && /^[12]:\d+$/.test(sourceId) ? null : ranges[index].id,
      };
    }),
  };
}
export function transformSectionState(
  before: Node,
  after: Node,
  state: WordSectionState,
  transactions: readonly Transaction[],
): WordSectionState {
  const oldParagraphs = sectionParagraphs(before),
    newParagraphs = sectionParagraphs(after);
  const endpoints = new Map(newParagraphs.map((p, i) => [p.to - 1, i]));
  const mapping = new Mapping();
  for (const tr of transactions) {
    const offset = mapping.maps.length;
    tr.mapping.maps.forEach((map, i) => {
      const step = tr.steps[i];
      const oldNode = step instanceof ReplaceAroundStep ? tr.docs[i].nodeAt(step.from) : null;
      const newNode =
        step instanceof ReplaceAroundStep ? (tr.docs[i + 1] || tr.doc).nodeAt(step.from) : null;
      // setNodeMarkup replaces paragraph wrapper tokens, including the token
      // carrying its section boundary. The paragraph and its content survive.
      // Treat that wrapper-only change as identity while retaining real joins.
      const wrapperOnly =
        step instanceof ReplaceAroundStep &&
        step.gapFrom === step.from + 1 &&
        step.gapTo === step.to - 1 &&
        step.insert === 1 &&
        oldNode?.isTextblock &&
        newNode?.isTextblock &&
        oldNode.nodeSize === step.to - step.from &&
        oldNode.nodeSize === newNode.nodeSize &&
        oldNode.content.eq(newNode.content);
      const mirror = tr.mapping.getMirror(i);
      mapping.appendMap(
        wrapperOnly ? StepMap.empty : map,
        mirror === undefined ? undefined : offset + mirror,
      );
    });
  }
  const breaks: WordSectionState['breaks'] = [];
  for (const boundary of state.breaks) {
    const p = oldParagraphs[boundary.paragraph];
    if (!p) throw Error('Missing live section paragraph.');
    const mapped = mapping.mapResult(p.to - 1, 1);
    if (mapped.deleted) continue;
    const paragraph = endpoints.get(mapped.pos);
    // A boundary inside a joined/replaced paragraph is gone, not moved to its far end.
    if (paragraph !== undefined) breaks.push({ sectionId: boundary.sectionId, paragraph });
  }
  if (state.version === 1) return { ...state, breaks };
  const ids = new Set([...breaks.map((b) => b.sectionId), state.finalSectionId]);
  return {
    ...state,
    breaks,
    inserted: state.inserted.filter((s) => ids.has(s.id)),
    starts: state.starts.filter((s) => ids.has(s.sectionId)),
    ...(state.layouts ? { layouts: state.layouts.filter((s) => ids.has(s.sectionId)) } : {}),
  };
}
