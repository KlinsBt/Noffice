import type { WordContent } from './model';
import { authoredWordSection } from './word-authored-section';
import { wordDefaults } from './word-defaults';
import { wordStoriesSchema } from './word-stories';

/** Fresh document stories inherit the application's explicit Normal font;
 * native Header/Footer override its paragraph spacing to Single and zero after.
 * Deriving these templates does not save metadata or create an undo event. */
export function authoredWordStoryTemplates(content: WordContent) {
  if (
    !authoredWordSection(
      content.sectionState?.version === 2 ? { ...content, sectionState: undefined } : content,
    )
  )
    return undefined;
  const stops = JSON.stringify([
    { position: 4536, alignment: 'center', leader: 'none' },
    { position: 9072, alignment: 'right', leader: 'none' },
  ]).replaceAll('"', '&quot;');
  const html = `<p data-source-paragraph="empty:0" data-word-tabs="${stops}" data-word-default-tab="720" style="font-family:${wordDefaults.fontFamily};font-size:${wordDefaults.fontSize}pt;line-height:1;margin-top:0pt;margin-bottom:0pt"></p>`;
  return { header: html, footer: html };
}

export function wordStoryCreationState(content: WordContent) {
  const templates = authoredWordStoryTemplates(content);
  return wordStoriesSchema.parse(
    templates
      ? {
          version: 1,
          evenAndOddHeaders: false,
          parts: [],
          ...content.stories,
          emptyTemplates: templates,
          templateVersion: 4,
        }
      : content.stories,
  );
}
