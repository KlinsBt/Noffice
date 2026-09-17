import type { WordContent } from './model';
import { clonedStoryHtml } from './word-story-provenance';
import { wordParagraphPointSpace } from './word-paragraph-spacing';

const selector = 'p,h1,h2,h3,h4,h5,h6';
const failure = () => Error('This saved Word document has ambiguous paragraph spacing after structural edits. Export a Noffice backup to preserve those edits, then reopen the original DOCX as a separate document.');
function restore(html: string, sourceHtml: string, legacyHtml: string) {
  const parse = (value: string) => new DOMParser().parseFromString(value, 'text/html');
  const current = parse(html), source = parse(sourceHtml), legacy = parse(legacyHtml);
  const paragraphs = [...current.querySelectorAll<HTMLElement>(selector)];
  const originals = [...source.querySelectorAll<HTMLElement>(selector)];
  const sources = new Map(originals.map(p => [p.dataset.sourceParagraph, p]));
  const old = new Map([...legacy.querySelectorAll<HTMLElement>(selector)].map(p => [p.dataset.sourceParagraph, p]));
  const alternate = originals.some(p => p.hasAttribute('data-word-space-before') || p.hasAttribute('data-word-space-after') || p.dataset.wordContextualSpacing === 'true');
  if (alternate && (originals.length !== paragraphs.length || sources.size !== originals.length ||
      originals.some((p, i) => !p.dataset.sourceParagraph || p.dataset.sourceParagraph !== paragraphs[i].dataset.sourceParagraph))) throw failure();
  for (const p of paragraphs) {
    const original = sources.get(p.dataset.sourceParagraph), previous = old.get(p.dataset.sourceParagraph);
    if (!original || !previous) continue;
    for (const attribute of ['data-word-paragraph-style', 'data-word-contextual-spacing'])
      if (!p.hasAttribute(attribute) && original.hasAttribute(attribute)) p.setAttribute(attribute, original.getAttribute(attribute)!);
    for (const [side, css] of [['before', 'marginTop'], ['after', 'marginBottom']] as const) {
      const attribute = 'data-word-space-' + side;
      if (p.hasAttribute(attribute) || !original.hasAttribute(attribute)) continue;
      // Old controls authored points only. Distinct edited point values remain
      // point intent; an unchanged cached value recovers the source mode.
      if (JSON.stringify(wordParagraphPointSpace(p.style[css])) !== JSON.stringify(wordParagraphPointSpace(previous.style[css]))) continue;
      p.setAttribute(attribute, original.getAttribute(attribute)!);
      p.style[css] = original.style[css];
    }
  }
  return current.body.innerHTML;
}

export function hydrateWordParagraphSpacing(current: WordContent, source: WordContent, legacy: WordContent): WordContent {
  const partHtml = (content: WordContent, part: NonNullable<WordContent['stories']>['parts'][number]) => {
    const original = content.stories?.parts.find(p => p.path === (part.copiedFrom || part.path));
    if (part.created) {
      const html = content.stories?.emptyTemplates?.[part.kind];
      return html ? clonedStoryHtml({ path: 'template', kind: part.kind, relationshipIds: [], html }, part.path) : null;
    }
    return original ? part.copiedFrom ? clonedStoryHtml(original, part.path) : original.html : null;
  };
  return { ...current, paragraphSpacingVersion: 1,
    html: restore(current.html, source.html, legacy.html),
    stories: current.stories && { ...current.stories, emptyTemplates: source.stories?.emptyTemplates,
      parts: current.stories.parts.map(part => {
        const original = partHtml(source, part), previous = partHtml(legacy, part);
        if (!original || !previous) throw failure();
        return { ...part, html: restore(part.html, original, previous) };
      }),
    },
  };
}
