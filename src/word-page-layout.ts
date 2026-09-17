import { Extension, type Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import type { WordContent } from './model';
import {
  initializeSectionState,
  validateSectionState,
  wordSectionSourceId,
  type AuthoredWordSectionState,
  type WordSectionState,
} from './word-section-breaks';
import { resolveWordSections } from './word-section-layout';
import {
  changedWordSectionGeometry,
  wordSectionGeometry,
  type WordSectionGeometryChange,
} from './word-section-geometry';

export type WordPageState = Pick<WordContent, 'paper' | 'margin' | 'orientation' | 'pageOverrides'>;
/** Editor-only document attribute: joins layout commands to ordinary text undo history.
 * It is saved as typed WordContent fields, never accepted from pasted HTML.
 */
export const WordPageLayout = Extension.create({
  name: 'wordPageLayout',
  addGlobalAttributes() {
    return [
      {
        types: ['doc'],
        attributes: { wordPageLayout: { default: null, rendered: false, parseHTML: () => null } },
      },
    ];
  },
});
export function initializeWordPageLayout(editor: Editor, content: WordContent) {
  const { paper, margin, orientation, pageOverrides } = content;
  editor.view.dispatch(
    editor.state.tr
      .setDocAttribute('wordPageLayout', { paper, margin, orientation, pageOverrides })
      .setMeta('addToHistory', false)
      .setMeta('preventUpdate', true),
  );
}
export function changeWordPageLayout<K extends 'paper' | 'margin' | 'orientation'>(
  editor: Editor,
  key: K,
  value: WordContent[K],
  content?: WordContent,
) {
  const current = editor.state.doc.attrs.wordPageLayout as WordPageState | null;
  if (!current) throw Error('Page layout is not initialized.');
  const next = {
    ...current,
    [key]: value,
    pageOverrides: { ...current.pageOverrides, [key]: true },
  };
  const sections = editor.state.doc.attrs.wordSectionState as WordSectionState | null;
  let updated = sections;
  if (sections?.version === 2 && sections.layouts?.length && !content)
    throw Error('Section-aware page setup requires the current document.');
  if (content && 'wordSectionState' in editor.state.doc.attrs) {
    const before =
      sections ||
      initializeSectionState(editor.state.doc, content.docxStructure) ||
      (!content.docxStructure
        ? {
            version: 2 as const,
            finalSectionId: 'authored-body',
            breaks: [],
            inserted: [],
            starts: [],
          }
        : null);
    const live = before
      ? resolveWordSections({
          ...content,
          ...current,
          html: editor.getHTML(),
          sectionState: before,
          stories: editor.state.doc.attrs.wordStories || content.stories,
        })
      : [];
    if (
      before &&
      live.length &&
      live.every(
        (section) =>
          section.simpleBody &&
          !section.columns &&
          section.margins.gutter === 0 &&
          !content.docxStructure?.sections
            .find((s) => s.id === wordSectionSourceId(before, section.id))
            ?.propertiesXml?.includes('sectPrChange'),
      )
    ) {
      validateSectionState(before, content.docxStructure, editor.state.doc);
      const authored: AuthoredWordSectionState =
        before.version === 2
          ? structuredClone(before)
          : { ...before, version: 2, inserted: [], starts: [] };
      const geometry = live.map(wordSectionGeometry);
      // Word applies the whole-document orientation setter in section order.
      // A boundary promoted by an earlier assignment remains Next Page even
      // when a later assignment makes the neighboring orientations equal.
      for (let index = 0; index < live.length; index++) {
        const old = geometry[index];
        const changed = changedWordSectionGeometry(old, {
          key,
          value,
        } as WordSectionGeometryChange);
        geometry[index] = changed;
        if (key !== 'orientation' || old.orientation === changed.orientation) continue;
        for (const affected of [
          index > 0 && geometry[index - 1].orientation !== changed.orientation
            ? live[index]
            : undefined,
          geometry[index + 1] && geometry[index + 1].orientation !== changed.orientation
            ? live[index + 1]
            : undefined,
        ]) {
          if (affected?.start !== 'continuous') continue;
          authored.starts = [
            ...authored.starts.filter((s) => s.sectionId !== affected.id),
            { sectionId: affected.id, start: 'nextPage' },
          ];
        }
      }
      authored.layouts = geometry;
      if (key === 'paper') {
        next.orientation = 'portrait';
        next.pageOverrides.orientation = true;
      }
      validateSectionState(authored, content.docxStructure, editor.state.doc);
      updated = authored;
    } else if (sections?.version === 2 && sections.layouts?.length) {
      throw Error('Page setup for this section layout is not supported yet.');
    }
  }
  if (
    JSON.stringify(next) === JSON.stringify(current) &&
    JSON.stringify(updated) === JSON.stringify(sections)
  )
    return;
  const tr = closeHistory(editor.state.tr).setDocAttribute('wordPageLayout', next);
  if (updated !== sections) tr.setDocAttribute('wordSectionState', updated);
  editor.view.dispatch(tr);
  // Keep the following typing operation out of this layout command's undo event.
  editor.view.dispatch(closeHistory(editor.state.tr));
  editor.view.focus();
}
