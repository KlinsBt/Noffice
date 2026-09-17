import type { Editor } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { closeHistory } from '@tiptap/pm/history';
import type { WordContent } from './model';
import { selectedWordSections } from './word-section-ranges';
import { resolveWordSections } from './word-section-layout';
import {
  initializeSectionState,
  mapLiveSections,
  validateSectionState,
  wordSectionSourceId,
  type AuthoredWordSectionState,
  type WordSectionState,
} from './word-section-breaks';
import {
  changedWordSectionGeometry,
  wordSectionGeometry,
  type WordSectionGeometryChange,
} from './word-section-geometry';
import { wordSectionStartSchema, type WordSectionStart } from './word-section-identity';

export type WordSectionLayoutChange =
  WordSectionGeometryChange | { key: 'start'; value: WordSectionStart };

/** A current-section command preserves body selection and joins metadata to the
 * ordinary document history. Cross-section selections need explicit scope. */
export function changeWordSectionLayout(
  editor: Editor,
  content: WordContent,
  change: WordSectionLayoutChange,
) {
  const { state } = editor;
  if (!(state.selection instanceof TextSelection))
    throw Error('Place the text cursor in the section to change.');
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
  if (!before) throw Error('The document sections cannot be mapped safely.');
  validateSectionState(before, content.docxStructure, state.doc);
  const ids = selectedWordSections(
    mapLiveSections(state.doc, content.docxStructure, before),
    state.selection.from,
    state.selection.to,
  );
  if (ids.length !== 1) throw Error('Select text within one section to change its page setup.');
  const current: WordContent = {
    ...content,
    ...(state.doc.attrs.wordPageLayout || {}),
    html: editor.getHTML(),
    sectionState: before,
    stories: state.doc.attrs.wordStories || content.stories,
  };
  const sections = resolveWordSections(current);
  const section = sections.find((s) => s.id === ids[0]);
  if (!section || !section.simpleBody || section.columns || section.margins.gutter !== 0)
    throw Error('Page setup for this section layout is not supported yet.');
  const original = content.docxStructure?.sections.find(
    (s) => s.id === wordSectionSourceId(before, section.id),
  );
  if (original?.propertiesXml?.includes('sectPrChange'))
    throw Error('Editing tracked section properties is not supported yet.');
  const next: AuthoredWordSectionState =
    before.version === 2
      ? structuredClone(before)
      : { ...before, version: 2, inserted: [], starts: [] };
  if (change.key === 'start') {
    const start = wordSectionStartSchema.parse(change.value);
    if (section.start === start) return false;
    next.starts = [
      ...next.starts.filter((s) => s.sectionId !== section.id),
      { sectionId: section.id, start },
    ];
  } else {
    const geometry = wordSectionGeometry(section);
    const changed = changedWordSectionGeometry(geometry, change);
    if (JSON.stringify(geometry) === JSON.stringify(changed)) return false;
    next.layouts = [...(next.layouts || []).filter((s) => s.sectionId !== section.id), changed];
    if (change.key === 'orientation' && changed.orientation !== section.orientation) {
      // Word's section-scoped orientation setter promotes adjacent continuous
      // boundaries to Next Page, in the same undo record as the geometry.
      const index = sections.indexOf(section);
      const previous = sections[index - 1],
        following = sections[index + 1];
      for (const affected of [
        previous && previous.orientation !== changed.orientation ? section : undefined,
        following && following.orientation !== changed.orientation ? following : undefined,
      ]) {
        if (affected?.start !== 'continuous') continue;
        next.starts = [
          ...next.starts.filter((s) => s.sectionId !== affected.id),
          { sectionId: affected.id, start: 'nextPage' },
        ];
      }
    }
  }
  validateSectionState(next, content.docxStructure, state.doc);
  editor.view.dispatch(closeHistory(state.tr).setDocAttribute('wordSectionState', next));
  editor.view.dispatch(closeHistory(editor.state.tr));
  editor.view.focus();
  return true;
}
