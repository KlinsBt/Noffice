import type { WordContent } from './model';
import { wordFontFeaturesValue } from './word-font-features';
import { clonedStoryHtml } from './word-story-provenance';

const failure = () => Error('This saved Word document has edited text with missing mixed font-feature metadata. Export a Noffice backup to keep those edits, then reopen the original DOCX as a separate document.');
const nodes = (element: HTMLElement) => {
  const walker = element.ownerDocument.createTreeWalker(element, NodeFilter.SHOW_TEXT), output: Text[] = [];
  while (walker.nextNode()) if (walker.currentNode.textContent) output.push(walker.currentNode as Text);
  return output;
};

function hydrateHtml(html: string, sourceHtml: string): string {
  const parse = (value: string) => new DOMParser().parseFromString(value, 'text/html');
  const current = parse(html), source = parse(sourceHtml);
  const sources = new Map([...source.querySelectorAll<HTMLElement>('[data-source-paragraph]')]
    .map((p) => [p.dataset.sourceParagraph, p]));
  for (const p of current.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6')) {
    const original = sources.get(p.dataset.sourceParagraph);
    if (!original) {
      if (p.textContent && p.querySelector('[data-word-font-features]') === null) throw failure();
      continue;
    }
    const mark = wordFontFeaturesValue(original.dataset.wordParagraphFontFeatures) ?? 0;
    if (wordFontFeaturesValue(p.dataset.wordParagraphFontFeatures) === null)
      p.dataset.wordParagraphFontFeatures = String(mark);
    const ranges: { end: number; value: number }[] = [];
    let offset = 0;
    for (const text of nodes(original)) {
      offset += text.length;
      const value = wordFontFeaturesValue(text.parentElement?.closest('[data-word-font-features]')
        ?.getAttribute('data-word-font-features')) ?? mark;
      if (ranges.at(-1)?.value === value) ranges.at(-1)!.end = offset;
      else ranges.push({ end: offset, value });
    }
    if (!ranges.length) ranges.push({ end: Infinity, value: mark });
    const uniform = ranges.length === 1;
    if (!uniform && p.textContent !== original.textContent) throw failure();
    let index = 0; offset = 0;
    for (const text of nodes(p)) {
      const end = offset + text.length;
      if (wordFontFeaturesValue(text.parentElement?.closest('[data-word-font-features]')
        ?.getAttribute('data-word-font-features')) !== null) { offset = end; continue; }
      const fragment = current.createDocumentFragment();
      let local = offset;
      while (local < end) {
        while (!uniform && index < ranges.length - 1 && ranges[index].end <= local) index++;
        const range = ranges[index], to = uniform ? end : Math.min(end, range.end);
        const span = current.createElement('span');
        span.dataset.wordFontFeatures = String(range.value);
        span.textContent = text.data.slice(local - offset, to - offset);
        fragment.append(span); local = to;
      }
      text.replaceWith(fragment); offset = end;
    }
  }
  return current.body.innerHTML;
}

/** Older editors could not change these switches. Uniform runs survive text
 * edits; unchanged mixed text has exact character identity. Never guess a
 * changed mixed run's settings or mutate the saved object before success. */
export function hydrateWordFontFeatures(current: WordContent, source: WordContent): WordContent {
  const html = hydrateHtml(current.html, source.html);
  const stories = current.stories ? { ...current.stories,
    emptyTemplates: source.stories?.emptyTemplates,
    parts: current.stories.parts.map((part) => {
      const original = source.stories?.parts.find((p) => p.path === (part.copiedFrom || part.path));
      let html = original?.html;
      if (part.created && source.stories?.emptyTemplates?.[part.kind])
        html = clonedStoryHtml({ path: 'template', kind: part.kind, relationshipIds: [],
          html: source.stories.emptyTemplates[part.kind]! }, part.path);
      else if (part.copiedFrom && original) html = clonedStoryHtml(original, part.path);
      if (!html) throw failure();
      return { ...part, html: hydrateHtml(part.html, html) };
    }),
  } : current.stories;
  return { ...current, html, stories, fontFeaturesVersion: 1 };
}
