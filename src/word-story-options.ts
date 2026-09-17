import type { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import type { WordContent } from './model';
import { wordStoriesSchema, type WordStories, wordStorySectionOptionsSchema } from './word-stories';
import { resolveWordSections } from './word-section-layout';

export interface WordStoryPageOptions {
  differentFirstPage: boolean;
  evenAndOddHeaders: boolean;
  headerDistance: number;
  footerDistance: number;
}

/** Native Single input rounds to a whole twip before range validation. The
 * captured half-twip cases round away from zero, including the zero boundary. */
export function wordStoryDistancePoints(input: string): number {
  const text = input.trim();
  if (text.length > 100 || !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(text)) return NaN;
  const points = Math.fround(Number(text));
  const value = Math.fround(points * 20);
  const twips = Math.sign(value) * Math.round(Math.abs(value));
  return Number.isFinite(twips) && twips >= 0 && twips <= 31680 ? twips || 0 : NaN;
}

/** Section overrides remain separate from immutable imported section XML.
 * The even-page switch belongs to the document; distances are in twips. */
export function changedWordStoryOptions(
  content: WordContent,
  sectionId: string,
  options: WordStoryPageOptions,
): WordStories {
  if (
    typeof options.differentFirstPage !== 'boolean' ||
    typeof options.evenAndOddHeaders !== 'boolean' ||
    [options.headerDistance, options.footerDistance].some(
      (n) => !Number.isInteger(n) || n < 0 || n > 31680,
    )
  )
    throw Error('Enter valid header and footer distances from 0 to 1,584 pt.');
  if (!resolveWordSections(content).some((s) => s.id === sectionId))
    throw Error('This section is no longer available.');
  const current = content.stories || { version: 1 as const, evenAndOddHeaders: false, parts: [] };
  const parsed = wordStorySectionOptionsSchema.parse({ sectionId, ...options });
  const source = resolveWordSections({
    ...content,
    stories: { ...current, sectionOptions: undefined },
  }).find((s) => s.id === sectionId)!;
  const override = {
    sectionId,
    ...(parsed.differentFirstPage !== source.differentFirstPage
      ? { differentFirstPage: parsed.differentFirstPage }
      : {}),
    ...(parsed.headerDistance !== source.margins.header
      ? { headerDistance: parsed.headerDistance }
      : {}),
    ...(parsed.footerDistance !== source.margins.footer
      ? { footerDistance: parsed.footerDistance }
      : {}),
  };
  const sectionOptions = (current.sectionOptions || []).filter((o) => o.sectionId !== sectionId);
  if (Object.keys(override).length > 1) sectionOptions.push(override);
  sectionOptions.sort((a, b) => a.sectionId.localeCompare(b.sectionId));
  const next: WordStories = { ...current, evenAndOddHeaders: options.evenAndOddHeaders };
  if (sectionOptions.length) next.sectionOptions = sectionOptions;
  else delete next.sectionOptions;
  return wordStoriesSchema.parse(next);
}

export function changeWordStoryPageOptions(
  editor: Editor,
  content: WordContent,
  expected: WordStories | undefined,
  sectionId: string,
  options: WordStoryPageOptions,
) {
  const current = editor.state.doc.attrs.wordStories as WordStories | null;
  if (JSON.stringify(current) !== JSON.stringify(expected || null))
    throw Error(
      'The headers or footers changed while these options were open. Reopen the options.',
    );
  const next = changedWordStoryOptions(
    { ...content, stories: current || undefined },
    sectionId,
    options,
  );
  if (JSON.stringify(next) === JSON.stringify(current)) return;
  if (!current && !next.evenAndOddHeaders && !next.sectionOptions?.length && !next.parts.length)
    return;
  editor.view.dispatch(closeHistory(editor.state.tr).setDocAttribute('wordStories', next));
  editor.view.dispatch(closeHistory(editor.state.tr));
}
