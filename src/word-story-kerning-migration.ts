import type { WordStories } from './word-stories';
import { hydrateParagraphMarkKerning } from './word-kerning-migration';
import { clonedStoryHtml } from './word-story-provenance';

/** Restore paragraph-mark metadata from retained source or generated provenance,
 * preserving edited story text and explicit run overrides. */
export function hydrateStoryParagraphKerning(
  current: WordStories,
  source: WordStories,
): WordStories {
  return {
    ...current,
    parts: current.parts.map((part) => {
      const original = source.parts.find((p) => p.path === (part.copiedFrom || part.path));
      let html = original?.html;
      if (part.created && source.emptyTemplates?.[part.kind])
        html = clonedStoryHtml(
          {
            path: 'template',
            kind: part.kind,
            relationshipIds: [],
            html: source.emptyTemplates[part.kind]!,
          },
          part.path,
        );
      else if (part.copiedFrom && original) html = clonedStoryHtml(original, part.path);
      if (!html)
        throw Error(
          'This saved header/footer has no source for its paragraph formatting. Keep a backup and reopen the original DOCX.',
        );
      const before = new DOMParser().parseFromString(html, 'text/html');
      const paragraphs = new Map(
        [...before.querySelectorAll<HTMLElement>('[data-source-paragraph]')].map((p) => [
          p.dataset.sourceParagraph,
          p,
        ]),
      );
      const next = new DOMParser().parseFromString(part.html, 'text/html');
      for (const paragraph of next.querySelectorAll<HTMLElement>('[data-source-paragraph]')) {
        const old = paragraphs.get(paragraph.dataset.sourceParagraph);
        if (old) hydrateParagraphMarkKerning(paragraph, old);
      }
      return { ...part, html: next.body.innerHTML };
    }),
  };
}
