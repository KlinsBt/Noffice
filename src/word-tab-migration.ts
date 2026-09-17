import type { WordContent } from './model';
import { clonedStoryHtml } from './word-story-provenance';

const settings = ['data-word-default-tab', 'data-word-decimal-symbol'] as const;
const paragraphAttributes = ['data-word-tabs', 'data-word-tab-layout-unsupported'] as const;

/** Older editors could insert tabs, but could not author their stop definitions.
 * Restore only absent metadata, retaining text, formatting and paragraph IDs. */
function restore(currentHtml: string, sourceHtml: string) {
  const source = new DOMParser().parseFromString(sourceHtml, 'text/html');
  const current = new DOMParser().parseFromString(currentHtml, 'text/html');
  const first = source.querySelector('p,h1,h2,h3,h4,h5,h6');
  const mapped = new Map([...source.querySelectorAll('[data-source-paragraph]')]
    .map((p) => [p.getAttribute('data-source-paragraph'), p]));
  for (const p of current.querySelectorAll('p,h1,h2,h3,h4,h5,h6')) {
    const original = mapped.get(p.getAttribute('data-source-paragraph'));
    for (const attribute of settings)
      if (!p.hasAttribute(attribute) && first?.hasAttribute(attribute))
        p.setAttribute(attribute, first.getAttribute(attribute)!);
    for (const attribute of paragraphAttributes)
      if (!p.hasAttribute(attribute) && original?.hasAttribute(attribute))
        p.setAttribute(attribute, original.getAttribute(attribute)!);
  }
  return current.body.innerHTML;
}

export function hydrateWordTabMetadata(current: WordContent, source: WordContent): WordContent {
  return {
    ...current,
    tabStopsVersion: 1,
    html: restore(current.html, source.html),
    stories: current.stories && {
      ...current.stories,
      parts: current.stories.parts.map((part) => {
        const original = source.stories?.parts.find((p) => p.path === (part.copiedFrom || part.path));
        let html = original?.html;
        if (part.created) {
          const template = source.stories?.emptyTemplates?.[part.kind];
          if (template) html = clonedStoryHtml({ path: 'template', kind: part.kind, relationshipIds: [], html: template }, part.path);
        } else if (part.copiedFrom && original) html = clonedStoryHtml(original, part.path);
        if (!html) throw Error('This saved header/footer has no source for its tab stops. Export a backup and reopen the original DOCX.');
        return { ...part, html: restore(part.html, html) };
      }),
    },
  };
}
