import type { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import type { WordContent } from './model';
import { wordStoriesSchema, type WordStories, type WordStoryReference } from './word-stories';
import { resolveWordSections } from './word-section-layout';

import { clonedStoryHtml } from './word-story-provenance';
export { clonedStoryHtml, clonedStoryParagraphKey } from './word-story-provenance';
import { materializeWordStoryRoots } from './word-story-create';

export function wordStoryIsLinked(
  content: WordContent,
  sectionId: string,
  kind: WordStoryReference['kind'],
  slot: WordStoryReference['slot'],
) {
  const sections = resolveWordSections(content),
    index = sections.findIndex((s) => s.id === sectionId);
  if (index < 1) return false;
  const override = content.stories?.references?.find(
    (r) => r.sectionId === sectionId && r.kind === kind && r.slot === slot,
  );
  if (override) return override.linked;
  const key = kind === 'header' ? 'headers' : 'footers';
  return (
    sections[index][key][slot]?.inherited ??
    !content.docxStructure?.sections
      .find((s) => s.id === sectionId)
      ?.[key].some((r) => r.type === slot)
  );
}

export function changedWordStoryLink(
  content: WordContent,
  sectionId: string,
  kind: WordStoryReference['kind'],
  slot: WordStoryReference['slot'],
  linked: boolean,
): WordStories {
  if (
    !['header', 'footer'].includes(kind) ||
    !['default', 'first', 'even'].includes(slot) ||
    typeof linked !== 'boolean'
  )
    throw Error('Choose a valid header or footer link.');
  const current = wordStoriesSchema.parse(content.stories);
  const sections = resolveWordSections(content),
    index = sections.findIndex((s) => s.id === sectionId);
  if (index < 1) throw Error('Only a later section can link to a previous header or footer.');
  const key = kind === 'header' ? 'headers' : 'footers';
  const selected = sections[index][key][slot];
  const wasLinked = wordStoryIsLinked(content, sectionId, kind, slot);
  if (linked === wasLinked) return current;
  let override: WordStoryReference;
  if (linked) override = { sectionId, kind, slot, linked: true };
  else {
    let part =
      selected && current.parts.find((p) => p.relationshipIds.includes(selected.relationshipId));
    if (!part) {
      materializeWordStoryRoots(content, current);
      const reference = resolveWordSections({ ...content, stories: current })[index][key][slot]!;
      part = current.parts.find((p) => p.relationshipIds.includes(reference.relationshipId))!;
    }
    const identity = crypto.randomUUID().replaceAll('-', '');
    const path = `word/noffice-${kind}-${identity}.xml`,
      relationshipId = `rIdNofficeStory${identity}`;
    current.parts.push({
      path,
      kind,
      relationshipIds: [relationshipId],
      ...(part.created ? { created: true as const } : { copiedFrom: part.copiedFrom || part.path }),
      html: clonedStoryHtml(part, path),
    });
    override = { sectionId, kind, slot, linked: false, relationshipId };
  }
  const references = (current.references || []).filter(
    (r) => !(r.sectionId === sectionId && r.kind === kind && r.slot === slot),
  );
  references.push(override);
  references.sort((a, b) =>
    `${a.sectionId}:${a.kind}:${a.slot}`.localeCompare(`${b.sectionId}:${b.kind}:${b.slot}`),
  );
  current.references = references;
  // Returning to an original inherited link removes its override and any unused
  // copied parts. Unreferenced original parts remain preserved in the package.
  const source = resolveWordSections({
    ...content,
    stories: { ...current, references: undefined },
  }).find((s) => s.id === sectionId)!;
  if (
    linked &&
    (source[key][slot]?.inherited ||
      (!source[key][slot] &&
        !content.docxStructure?.sections
          .find((s) => s.id === sectionId)
          ?.[key].some((r) => r.type === slot)))
  ) {
    current.references = references.filter((r) => r !== override);
  }
  if (!current.references.length) delete current.references;
  const used = new Set(
    resolveWordSections({ ...content, stories: current }).flatMap((s) =>
      [...Object.values(s.headers), ...Object.values(s.footers)].flatMap((r) =>
        r ? [r.relationshipId] : [],
      ),
    ),
  );
  current.parts = current.parts.filter(
    (p) => (!p.copiedFrom && !p.created) || p.relationshipIds.some((id) => used.has(id)),
  );
  return wordStoriesSchema.parse(current);
}

export function changeWordStoryLink(
  editor: Editor,
  content: WordContent,
  expected: WordStories | undefined,
  sectionId: string,
  kind: WordStoryReference['kind'],
  slot: WordStoryReference['slot'],
  linked: boolean,
) {
  const current = editor.state.doc.attrs.wordStories as WordStories | null;
  if (JSON.stringify(current) !== JSON.stringify(expected || null))
    throw Error('The headers or footers changed. Reopen the link controls.');
  const next = changedWordStoryLink(
    { ...content, stories: current || undefined },
    sectionId,
    kind,
    slot,
    linked,
  );
  if (JSON.stringify(next) === JSON.stringify(current)) return;
  editor.view.dispatch(closeHistory(editor.state.tr).setDocAttribute('wordStories', next));
  editor.view.dispatch(closeHistory(editor.state.tr));
}
