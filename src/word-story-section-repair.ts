import { emptyWordStory } from './word-story-create';
import type { DocxStructure } from './docx-sections';
import type { WordContent } from './model';
import type { WordSectionState } from './word-section-breaks';
import { resolveWordSections } from './word-section-layout';
import { wordStoriesSchema, type WordStories, type WordStoryReference } from './word-stories';

/** Native section deletion preserves a surviving story's effective contents,
 * materializing its reference when the previous owner disappears. Repair the
 * same document transaction, including copied parts and per-section options. */
export function repairWordStorySections(
  source: DocxStructure | undefined,
  stories: WordStories,
  before: WordSectionState,
  after: WordSectionState,
): WordStories {
  const ids = (state: WordSectionState) => [
    ...state.breaks.map((b) => b.sectionId),
    state.finalSectionId,
  ];
  const oldIds = new Set(ids(before)),
    live = new Set(ids(after));
  if (live.size >= oldIds.size || [...live].some((id) => !oldIds.has(id))) return stories;
  // Only reference identities are consumed here; global paper fields do not
  // participate in this repair and never enter the persisted result.
  const content: WordContent = {
    kind: 'word',
    html: '',
    paper: 'a4',
    margin: 'normal',
    docxStructure: source,
    stories,
    sectionState: before,
  };
  const all = resolveWordSections(content);
  const old = all.filter((s) => live.has(s.id));
  const firstIndex = all.findIndex((s) => live.has(s.id));
  // Word transfers the nearest story-owning section as a group when leading
  // boundaries disappear. Slots inherited from an earlier owner become blank.
  const owner =
    firstIndex > 0
      ? all
          .slice(0, firstIndex + 1)
          .findLast((s) =>
            [...Object.values(s.headers), ...Object.values(s.footers)].some(
              (r) => r && !r.inherited,
            ),
          )
      : undefined;
  const reset = new Map<
    string,
    { relationshipId: string; sourceSectionId: string; inherited: boolean }
  >();
  const next = structuredClone(stories);
  delete next.references;
  if (next.sectionOptions) {
    next.sectionOptions = next.sectionOptions.filter((o) => live.has(o.sectionId));
    if (!next.sectionOptions.length) delete next.sectionOptions;
  }
  const baseline = resolveWordSections({ ...content, stories: next, sectionState: after });
  const overrides: WordStoryReference[] = [];
  const used = new Set<string>();
  let previous: (typeof old)[number] | undefined;
  for (const [index, section] of old.entries()) {
    for (const [kind, key] of [
      ['header', 'headers'],
      ['footer', 'footers'],
    ] as const) {
      for (const slot of ['default', 'first', 'even'] as const) {
        const inherited = previous?.[key][slot];
        let wanted = section[key][slot],
          actual = baseline[index][key][slot];
        const resetKey = `${kind}:${slot}`;
        if (
          index === 0 &&
          owner &&
          wanted &&
          all.findIndex((s) => s.id === wanted!.sourceSectionId) < all.indexOf(owner)
        ) {
          const part = emptyWordStory(next, kind);
          next.parts.push(part);
          wanted = {
            relationshipId: part.relationshipIds[0],
            sourceSectionId: section.id,
            inherited: false,
          };
          reset.set(resetKey, wanted);
        } else if (wanted?.inherited && reset.has(resetKey)) {
          wanted = { ...reset.get(resetKey)!, inherited: true };
        } else reset.delete(resetKey);
        if (wanted) {
          wanted = {
            ...wanted,
            inherited: !!wanted.inherited && inherited?.relationshipId === wanted.relationshipId,
          };
          used.add(wanted.relationshipId);
        }
        if (actual?.inherited) actual = inherited ? { ...inherited, inherited: true } : null;
        if (
          actual?.relationshipId !== wanted?.relationshipId ||
          !!actual?.inherited !== !!wanted?.inherited
        ) {
          overrides.push(
            wanted && !wanted.inherited
              ? {
                  sectionId: section.id,
                  kind,
                  slot,
                  linked: false,
                  relationshipId: wanted.relationshipId,
                }
              : { sectionId: section.id, kind, slot, linked: true },
          );
        }
        section[key][slot] = wanted;
      }
    }
    previous = section;
  }
  if (overrides.length) next.references = overrides;
  next.parts = next.parts.filter(
    (p) => (!p.copiedFrom && !p.created) || p.relationshipIds.some((id) => used.has(id)),
  );
  return wordStoriesSchema.parse(next);
}
