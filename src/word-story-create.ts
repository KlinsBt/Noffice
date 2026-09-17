import type { WordStories } from './word-stories';
import { clonedStoryHtml } from './word-story-provenance';
import { resolveWordSections } from './word-section-layout';
import type { WordContent } from './model';
import type { WordStoryReference } from './word-stories';
import { wordStoriesSchema } from './word-stories';
import { sanitizeHTML } from './formats';
import { wordJSON } from './word-extensions';
import { closeHistory } from '@tiptap/pm/history';
import type { Editor } from '@tiptap/core';
import { wordStoryCreationState } from './word-authored-stories';

/** Create a private styled blank part; the caller commits its reference and
 * body change in one document history transaction. */
export function emptyWordStory(stories: WordStories, kind: 'header' | 'footer') {
  const html = stories.emptyTemplates?.[kind];
  if (!html)
    throw Error(
      'Reopen the original DOCX to restore its header/footer style before creating an empty story.',
    );
  const identity = crypto.randomUUID().replaceAll('-', '');
  const part = {
    path: `word/noffice-${kind}-${identity}.xml`,
    kind,
    relationshipIds: [`rIdNofficeStory${identity}`],
    created: true as const,
    html,
  };
  part.html = clonedStoryHtml({ ...part, created: undefined }, part.path);
  return part;
}

/** Word materializes all six first-section slots when a story is authored or
 * explicitly unlinked, including inactive first/even slots. Preserve all
 * existing parts and references; later missing slots continue to inherit. */
export function materializeWordStoryRoots(content: WordContent, stories: WordStories) {
  const first = resolveWordSections({ ...content, stories })[0];
  if (!first) throw Error('The document has no editable section for this story.');
  for (const [kind, key] of [
    ['header', 'headers'],
    ['footer', 'footers'],
  ] as const)
    for (const slot of ['default', 'first', 'even'] as const) {
      if (first[key][slot]) continue;
      const part = emptyWordStory(stories, kind);
      stories.parts.push(part);
      stories.references = (stories.references || []).filter(
        (r) => !(r.sectionId === first.id && r.kind === kind && r.slot === slot),
      );
      stories.references.push({
        sectionId: first.id,
        kind,
        slot,
        linked: false,
        relationshipId: part.relationshipIds[0],
      });
    }
}

export function changedWordStoryCreation(
  content: WordContent,
  sectionId: string,
  kind: WordStoryReference['kind'],
  slot: WordStoryReference['slot'],
  html: string,
) {
  if (!['header', 'footer'].includes(kind) || !['default', 'first', 'even'].includes(slot))
    throw Error('Choose a valid header or footer.');
  if (html.length > 1000000) throw Error('This header or footer exceeds the supported size.');
  const stories = wordStoryCreationState(content);
  const sections = resolveWordSections(content),
    section = sections.find((s) => s.id === sectionId);
  if (!section) throw Error('This section changed. Reopen the header or footer.');
  const key = kind === 'header' ? 'headers' : 'footers';
  if (section[key][slot])
    throw Error('This header or footer already exists. Reopen it before editing.');
  const template = stories.emptyTemplates?.[kind];
  if (!template)
    throw Error('The document is missing the styles needed to create this header or footer.');
  const clean = sanitizeHTML(html);
  // Opening and closing an untouched blank story does not dirty the document.
  if (JSON.stringify(wordJSON(clean)) === JSON.stringify(wordJSON(template))) return stories;
  const dom = new DOMParser().parseFromString(clean, 'text/html');
  if (
    [...dom.querySelectorAll('[data-source-paragraph]')].some(
      (p) => p.getAttribute('data-source-paragraph') !== 'empty:0',
    )
  )
    throw Error('The new header/footer has invalid paragraph provenance.');
  materializeWordStoryRoots(content, stories);
  const ref = resolveWordSections({ ...content, stories }).find((s) => s.id === sectionId)![key][
    slot
  ]!;
  const part = stories.parts.find((p) => p.relationshipIds.includes(ref.relationshipId))!;
  part.html = clonedStoryHtml({ ...part, created: undefined, html: clean }, part.path);
  return wordStoriesSchema.parse(stories);
}

export function createWordStory(
  editor: Editor,
  content: WordContent,
  expected: WordStories | undefined,
  sectionId: string,
  kind: WordStoryReference['kind'],
  slot: WordStoryReference['slot'],
  html: string,
) {
  const current = editor.state.doc.attrs.wordStories as WordStories | null;
  if (JSON.stringify(current) !== JSON.stringify(expected || null))
    throw Error('The headers or footers changed. Reopen the story controls.');
  const next = changedWordStoryCreation(
    { ...content, stories: current || undefined },
    sectionId,
    kind,
    slot,
    html,
  );
  if (JSON.stringify(next) === JSON.stringify(current)) return;
  // A fresh document has no persisted story state until a real change. Merely
  // opening and applying its derived blank template must remain a no-op.
  if (next.parts.length === (current?.parts.length || 0)) return;
  editor.view.dispatch(closeHistory(editor.state.tr).setDocAttribute('wordStories', next));
  editor.view.dispatch(closeHistory(editor.state.tr));
}
