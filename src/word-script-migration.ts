import type { WordContent } from './model';
import { wordScriptValue } from './word-script';
import { clonedStoryHtml } from './word-story-provenance';

const failure = () => Error('This saved Word document has changed paragraph structure with missing script formatting. Export a Noffice backup to keep those edits, then reopen the original DOCX as a separate document.');

function hydrateHtml(html: string, sourceHtml: string): string {
  const parse = (value: string) => new DOMParser().parseFromString(value, 'text/html');
  const current = parse(html), source = parse(sourceHtml);
  const paragraphs = [...current.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6')];
  const originals = [...source.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6')];
  const baseline = originals.every(p => (wordScriptValue(p.dataset.wordParagraphScript) || 'baseline') === 'baseline');
  const sources = new Map(originals.map(p => [p.dataset.sourceParagraph, p]));
  const sameStructure = originals.length === paragraphs.length && originals.every((p, i) =>
    p.dataset.sourceParagraph && p.dataset.sourceParagraph === paragraphs[i].dataset.sourceParagraph) && sources.size === originals.length;
  for (const p of paragraphs) {
    if (wordScriptValue(p.dataset.wordParagraphScript)) continue;
    // Old editors could change inline scripts, but never paragraph scripts.
    // A split/join may have changed which source mark survived. Its identity
    // cannot be recovered from text alone when the source has script marks.
    if (!baseline && !sameStructure) throw failure();
    const original = sources.get(p.dataset.sourceParagraph);
    p.dataset.wordParagraphScript = baseline ? 'baseline'
      : wordScriptValue(original?.dataset.wordParagraphScript) || 'baseline';
  }
  return current.body.innerHTML;
}

/** Recover mark metadata without reinterpreting or replacing saved inline edits. */
export function hydrateWordParagraphScripts(current: WordContent, source: WordContent): WordContent {
  const html = hydrateHtml(current.html, source.html);
  const stories = current.stories ? { ...current.stories,
    emptyTemplates: source.stories?.emptyTemplates,
    parts: current.stories.parts.map(part => {
      const original = source.stories?.parts.find(p => p.path === (part.copiedFrom || part.path));
      let html = original?.html;
      if (part.created && source.stories?.emptyTemplates?.[part.kind])
        html = clonedStoryHtml({ path: 'template', kind: part.kind, relationshipIds: [],
          html: source.stories.emptyTemplates[part.kind]! }, part.path);
      else if (part.copiedFrom && original) html = clonedStoryHtml(original, part.path);
      if (!html) throw failure();
      return { ...part, html: hydrateHtml(part.html, html) };
    }),
  } : current.stories;
  return { ...current, html, stories, paragraphScriptsVersion: 1 };
}
