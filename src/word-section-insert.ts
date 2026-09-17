import type { Editor } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { closeHistory } from '@tiptap/pm/history';
import { canSplit } from '@tiptap/pm/transform';
import type { WordContent } from './model';
import { resolveWordSections } from './word-section-layout';
import { wordSectionStartSchema, type WordSectionStart } from './word-section-identity';
import {
  sectionParagraphs,
  initializeSectionState,
  mapLiveSections,
  transformSectionState,
  validateSectionState,
  wordSectionSourceId,
  type AuthoredWordSectionState,
  type WordSectionState,
} from './word-section-breaks';
import { wordStoriesSchema } from './word-stories';
import { wordSectionGeometry } from './word-section-geometry';
import { inFlatWordList } from './word-flat-list';

/** Word inserts before a selected range without replacing its text. The left
 * section takes a new identity; the following section owns the requested start.
 * Paragraphs, section metadata and story links share one history transaction. */
export function insertWordSectionBreak(
  editor: Editor,
  content: WordContent,
  start: WordSectionStart,
) {
  wordSectionStartSchema.parse(start);
  const { state } = editor,
    { selection } = state;
  const splitDepth = inFlatWordList(selection.$from) ? 2 : 1;
  if (
    !(selection instanceof TextSelection) ||
    (selection.$from.depth !== 1 && splitDepth !== 2) ||
    !['paragraph', 'heading'].includes(selection.$from.parent.type.name) ||
    !canSplit(state.doc, selection.from, splitDepth)
  )
    return false;
  const before =
    (state.doc.attrs.wordSectionState as WordSectionState | null) ||
    initializeSectionState(state.doc, content.docxStructure) ||
    (!content.docxStructure
      ? {
          version: 2 as const,
          finalSectionId: 'authored-body',
          breaks: [],
          inserted: [],
          starts: [],
        }
      : null);
  if (!before) return false;
  validateSectionState(before, content.docxStructure, state.doc);
  if (before.breaks.length >= 9999) throw Error('The document has reached the section limit.');
  const current: WordContent = {
    ...content,
    html: editor.getHTML(),
    sectionState: before,
    stories: state.doc.attrs.wordStories || content.stories,
  };
  const ranges = mapLiveSections(state.doc, content.docxStructure, before);
  const range = ranges.ranges.find((r) => selection.from >= r.from && selection.from < r.to);
  const section = resolveWordSections(current).find((s) => s.id === range?.id);
  if (!section || !wordSectionStartSchema.safeParse(section.start).success) return false;
  const sourceId = wordSectionSourceId(before, section.id);
  const original = content.docxStructure?.sections.find((s) => s.id === sourceId);
  if (original?.propertiesXml?.includes('sectPrChange'))
    throw Error('Inserting sections with tracked section properties is not supported yet.');
  const id = `authored-section:${crypto.randomUUID().replaceAll('-', '')}`;
  const tr = closeHistory(state.tr).split(selection.from, splitDepth);
  const mapped = transformSectionState(state.doc, tr.doc, before, [tr]);
  const paragraphs = sectionParagraphs(tr.doc);
  const paragraph = paragraphs.findIndex((p) => p.to - 1 === selection.from);
  if (paragraph < 0) throw Error('The inserted section has no paragraph boundary.');
  const next: AuthoredWordSectionState = {
    ...mapped,
    version: 2,
    inserted: [...(before.version === 2 ? before.inserted : []), { id, sourceSectionId: sourceId }],
    starts: [
      ...(before.version === 2 ? before.starts.filter((s) => s.sectionId !== section.id) : []),
      { sectionId: id, start: section.start as WordSectionStart },
      { sectionId: section.id, start },
    ],
    breaks: [...mapped.breaks, { sectionId: id, paragraph }].sort(
      (a, b) => a.paragraph - b.paragraph,
    ),
    ...(before.version === 2 && before.layouts?.some((s) => s.sectionId === section.id)
      ? {
          layouts: [
            ...(mapped.version === 2 ? mapped.layouts || [] : []),
            wordSectionGeometry({ ...section, id }),
          ],
        }
      : {}),
  };
  validateSectionState(next, content.docxStructure, tr.doc);
  tr.setDocAttribute('wordSectionState', next);
  if (current.stories) {
    const stories = structuredClone(current.stories);
    stories.sectionOptions = [
      ...(stories.sectionOptions || []),
      {
        sectionId: id,
        differentFirstPage: section.differentFirstPage,
        ...(section.margins.header !== null ? { headerDistance: section.margins.header } : {}),
        ...(section.margins.footer !== null ? { footerDistance: section.margins.footer } : {}),
      },
    ];
    stories.references = (stories.references || []).filter((r) => r.sectionId !== section.id);
    for (const [kind, key] of [
      ['header', 'headers'],
      ['footer', 'footers'],
    ] as const) {
      for (const slot of ['default', 'first', 'even'] as const) {
        const ref = section[key][slot];
        stories.references.push(
          ref && !ref.inherited
            ? { sectionId: id, kind, slot, linked: false, relationshipId: ref.relationshipId }
            : { sectionId: id, kind, slot, linked: true },
        );
        stories.references.push({ sectionId: section.id, kind, slot, linked: true });
      }
    }
    tr.setDocAttribute('wordStories', wordStoriesSchema.parse(stories));
  }
  tr.setSelection(TextSelection.create(tr.doc, selection.from + splitDepth * 2));
  editor.view.dispatch(tr.scrollIntoView());
  editor.view.dispatch(closeHistory(editor.state.tr));
  editor.view.focus();
  return true;
}
