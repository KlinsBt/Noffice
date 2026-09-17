import type { WordContent } from './model';
import { assertXmlComplexity } from './office-preservation';
import { authoredWordSection } from './word-authored-section';
import { validateWordSectionIdentities, wordSectionSourceId } from './word-section-breaks';
import { wordFinalSectionProperties } from './word-section-defaults';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const slots = ['default', 'even', 'first'] as const;
type Slot = (typeof slots)[number];
type Reference = { relationshipId: string; sourceSectionId: string; inherited: boolean };
type References = Record<Slot, Reference | null>;
const sides = ['left', 'right', 'top', 'bottom', 'header', 'footer', 'gutter'] as const;
type Side = (typeof sides)[number];
export interface WordSectionLayout {
  id: string;
  paragraphs: string[];
  width: number | null;
  height: number | null;
  orientation: 'portrait' | 'landscape';
  margins: Record<Side, number | null>;
  headers: References;
  footers: References;
  differentFirstPage: boolean;
  simpleBody: boolean;
  start: string;
  columns?: { widths: number[]; spaces: number[]; separator: boolean };
}
const child = (e: Element | null, name: string) =>
  e ? [...e.children].find((c) => c.namespaceURI === W && c.localName === name) : undefined;
const val = (e: Element | undefined, name = 'val') => e?.getAttributeNS(W, name) ?? '';
const on = (e: Element | undefined) => !!e && !['0', 'false', 'off'].includes(val(e));
function twips(e: Element | undefined, name: string, signed = false): number | null {
  const raw = val(e, name);
  if (!/^[+-]?\d+$/.test(raw)) return null;
  const n = Number(raw);
  return Number.isSafeInteger(n) && n >= (signed ? -31680 : 0) && n <= 31680 ? n : null;
}
function properties(xml: string | null) {
  if (!xml) return null;
  if (xml.length > 100000) throw Error('Section properties exceed the supported size.');
  assertXmlComplexity(xml);
  const d = new DOMParser().parseFromString(xml, 'application/xml');
  if (
    d.getElementsByTagName('parsererror').length ||
    d.documentElement.namespaceURI !== W ||
    d.documentElement.localName !== 'sectPr'
  )
    throw Error('The Word section properties contain invalid XML.');
  return d.documentElement;
}

/** Resolve retained source properties without mistaking historical XML for live state.
 * Null measurements are deliberately unknown, not assumed native defaults.
 * Header/footer slots resolve independently; page selection also needs document settings.
 */
export function resolveWordSections(content: WordContent): WordSectionLayout[] {
  const sections = content.docxStructure?.sections;
  if (!sections?.length) {
    const authored = authoredWordSection({
      ...content,
      sectionState: undefined,
      stories:
        content.sectionState?.version === 2 && content.stories
          ? { ...content.stories, sectionOptions: undefined }
          : content.stories,
    });
    return authored ? resolveLiveWordSections(content, [authored]) : [];
  }
  const roots = sections.map((s) => {
    const root = properties(s.propertiesXml);
    return s.endingParagraph === null
      ? wordFinalSectionProperties(root, content.docxStructure?.compatibility?.mode)
      : root;
  });
  const last = roots.at(-1)!,
    size = child(last, 'pgSz'),
    margin = child(last, 'pgMar');
  // Match the legacy global controls/import baseline. Section-scoped commands will
  // need explicit overrides, rather than changing immutable source properties.
  const baseline = {
    paper: [Number(val(size, 'w')), Number(val(size, 'h'))].includes(12240) ? 'letter' : 'a4',
    orientation: val(size, 'orient') === 'landscape' ? 'landscape' : 'portrait',
    margin:
      val(margin, 'left') === '720' ? 'narrow' : val(margin, 'left') === '2160' ? 'wide' : 'normal',
  };
  let previousHeaders: References = { default: null, even: null, first: null };
  let previousFooters: References = { default: null, even: null, first: null };
  const resolved: WordSectionLayout[] = sections.map((source, i) => {
    const root = roots[i],
      size = child(root, 'pgSz'),
      margin = child(root, 'pgMar');
    let width = twips(size, 'w'),
      height = twips(size, 'h');
    if (!width) width = null;
    if (!height) height = null;
    let orientation: WordSectionLayout['orientation'] =
      val(size, 'orient') === 'landscape' ? 'landscape' : 'portrait';
    // Authored sections clone immutable source defaults. Live options are
    // applied to each identity below, never inherited from a sibling's edits.
    const options =
      content.sectionState?.version === 2
        ? undefined
        : content.stories?.sectionOptions?.find((o) => o.sectionId === source.id);
    const margins = Object.fromEntries(
      sides.map((side) => [side, twips(margin, side, side === 'top' || side === 'bottom')]),
    ) as WordSectionLayout['margins'];
    if (options?.headerDistance !== undefined) margins.header = options.headerDistance;
    if (options?.footerDistance !== undefined) margins.footer = options.footerDistance;
    margins.gutter ??= 0;
    if (
      content.pageOverrides?.paper ||
      content.pageOverrides?.orientation ||
      content.paper !== baseline.paper ||
      (content.orientation || 'portrait') !== baseline.orientation
    ) {
      let dimensions = [width, height];
      if (
        content.pageOverrides?.paper ||
        content.paper !== baseline.paper ||
        width === null ||
        height === null
      )
        dimensions = content.paper === 'letter' ? [12240, 15840] : [11906, 16838];
      const sorted = (dimensions as number[]).sort((a, b) => a - b);
      orientation = content.orientation || 'portrait';
      if (orientation === 'landscape') sorted.reverse();
      [width, height] = sorted;
    }
    if (content.pageOverrides?.margin || content.margin !== baseline.margin) {
      const value = content.margin === 'narrow' ? 720 : content.margin === 'wide' ? 2160 : 1440;
      for (const side of ['left', 'right', 'top', 'bottom'] as const) margins[side] = value;
    }
    const references = (kind: 'headers' | 'footers', previous: References) =>
      Object.fromEntries(
        slots.map((slot) => {
          const direct = source[kind].find((r) => r.type === slot);
          return [
            slot,
            direct
              ? {
                  relationshipId: direct.relationshipId,
                  sourceSectionId: source.id,
                  inherited: false,
                }
              : previous[slot]
                ? { ...previous[slot], inherited: true }
                : null,
          ];
        }),
      ) as References;
    const headers = references('headers', previousHeaders),
      footers = references('footers', previousFooters);
    previousHeaders = headers;
    previousFooters = footers;
    const columns = child(root, 'cols');
    const count = Number(val(columns, 'num') || '1');
    const equal = !['0', 'false', 'off'].includes(val(columns, 'equalWidth'));
    const bodyWidth =
      width === null || margins.left === null || margins.right === null
        ? null
        : width - margins.left - margins.right;
    const gap = columns?.hasAttributeNS(W, 'space') ? twips(columns, 'space') : 720;
    const explicit = columns
      ? [...columns.children].filter((c) => c.namespaceURI === W && c.localName === 'col')
      : [];
    const widths =
      equal && bodyWidth !== null && gap !== null
        ? Array.from(
            { length: Number.isSafeInteger(count) && count > 0 && count <= 64 ? count : 0 },
            () => (bodyWidth - gap * (count - 1)) / count,
          )
        : explicit.map((c) => twips(c, 'w'));
    const spaces = equal
      ? widths.map((_, i) => (i === widths.length - 1 ? 0 : gap))
      : explicit.map((c, i) => (i === explicit.length - 1 ? 0 : (twips(c, 'space') ?? gap)));
    const validColumns =
      Number.isSafeInteger(count) &&
      count > 0 &&
      count <= 64 &&
      widths.length === count &&
      widths.every((w) => w !== null && w > 0) &&
      spaces.every((s) => s !== null && s >= 0) &&
      bodyWidth !== null &&
      Math.abs(
        widths.reduce<number>((n, w) => n + (w ?? 0), 0) +
          spaces.reduce<number>((n, s) => n + (s ?? 0), 0) -
          bodyWidth,
      ) < 1;
    return {
      id: source.id,
      paragraphs: [...source.paragraphs],
      width,
      height,
      orientation,
      margins,
      headers,
      footers,
      differentFirstPage: options?.differentFirstPage ?? on(child(root, 'titlePg')),
      start: val(child(root, 'type')) || 'nextPage',
      ...(validColumns && count > 1
        ? {
            columns: {
              widths: widths as number[],
              spaces: spaces as number[],
              separator:
                on(child(root, 'cols')) && !['', '0', 'false', 'off'].includes(val(columns, 'sep')),
            },
          }
        : {}),
      simpleBody:
        (!columns || validColumns) &&
        !on(child(root, 'bidi')) &&
        !on(child(root, 'rtlGutter')) &&
        ['', 'top'].includes(val(child(root, 'vAlign'))),
    };
  });
  return resolveLiveWordSections(content, resolved);
}

function resolveLiveWordSections(
  content: WordContent,
  resolved: WordSectionLayout[],
): WordSectionLayout[] {
  const state = content.sectionState;
  if (state) validateWordSectionIdentities(state, content.docxStructure);
  const liveIds =
    content.sectionState &&
    new Set([
      ...content.sectionState.breaks.map((b) => b.sectionId),
      content.sectionState.finalSectionId,
    ]);
  const surviving =
    state?.version === 2
      ? [...state.breaks.map((b) => b.sectionId), state.finalSectionId].map((id) => {
          const source = resolved.find((s) => s.id === wordSectionSourceId(state, id));
          if (!source) throw Error('Missing authored section source.');
          const section = {
            ...source,
            id,
            paragraphs: [...source.paragraphs],
            margins: { ...source.margins },
            headers: { ...source.headers },
            footers: { ...source.footers },
          };
          const options = content.stories?.sectionOptions?.find((o) => o.sectionId === id);
          if (options?.differentFirstPage !== undefined)
            section.differentFirstPage = options.differentFirstPage;
          if (options?.headerDistance !== undefined)
            section.margins.header = options.headerDistance;
          if (options?.footerDistance !== undefined)
            section.margins.footer = options.footerDistance;
          section.start = state.starts.find((s) => s.sectionId === id)?.start ?? section.start;
          const layout = state.layouts?.find((s) => s.sectionId === id);
          if (layout) {
            section.width = layout.width;
            section.height = layout.height;
            section.orientation = layout.orientation;
            Object.assign(section.margins, layout.margins);
            // Column redistribution needs its own native authoring contract.
            if (section.columns) section.simpleBody = false;
          }
          return section;
        })
      : liveIds
        ? resolved.filter((s) => liveIds.has(s.id))
        : resolved;
  let previous: WordSectionLayout | undefined;
  for (const section of surviving) {
    for (const kind of ['headers', 'footers'] as const) {
      for (const slot of slots) {
        const reference = section[kind][slot];
        // The writer materializes an inherited reference if its former predecessor
        // disappeared. Preserve its source provenance, but report the current link.
        if (
          reference?.inherited &&
          previous?.[kind][slot]?.relationshipId !== reference.relationshipId
        )
          section[kind][slot] = { ...reference, inherited: false };
      }
    }
    previous = section;
  }
  // First preserve links across deleted source sections. Then apply live commands
  // in surviving section order, so inherited successors follow a newly unlinked part.
  const overrides = new Map(
    (content.stories?.references || []).map((r) => [`${r.sectionId}:${r.kind}:${r.slot}`, r]),
  );
  previous = undefined;
  for (const section of surviving) {
    for (const kind of ['headers', 'footers'] as const)
      for (const slot of slots) {
        const override = overrides.get(
          `${section.id}:${kind === 'headers' ? 'header' : 'footer'}:${slot}`,
        );
        if (override && !override.linked) {
          section[kind][slot] = {
            relationshipId: override.relationshipId,
            sourceSectionId: section.id,
            inherited: false,
          };
        } else if (override?.linked || section[kind][slot]?.inherited || !section[kind][slot]) {
          // A previously empty slot inherits a newly created predecessor too;
          // inheritance is not conditional on an original package part existing.
          const prior = previous?.[kind][slot];
          section[kind][slot] = prior ? { ...prior, inherited: true } : null;
        }
      }
    previous = section;
  }
  return surviving;
}

/** Rectangle supported by the current single-surface editor, in twentieths of a point. */
export function singleSectionPage(content: WordContent): WordSectionLayout | null {
  const authored = authoredWordSection(content);
  if (authored) return authored;
  if (!content.sectionState && content.docxStructure?.sections.length !== 1) return null;
  const sections = resolveWordSections(content);
  if (sections.length !== 1) return null;
  const s = sections[0];
  if (!s.simpleBody || s.columns || !s.width || !s.height || s.margins.gutter !== 0) return null;
  const { left, right, top, bottom } = s.margins;
  if (
    [left, right, top, bottom].some((n) => n === null || n < 0) ||
    left! + right! >= s.width ||
    top! + bottom! >= s.height
  )
    return null;
  return s;
}
