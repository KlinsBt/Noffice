import type { WordContent } from './model';
import { clonedStoryHtml } from './word-story-provenance';
import { wordHyphenText } from './word-hyphen';

const failure = () => Error('This saved Word document has ambiguous edited hyphen characters. Export a Noffice backup to preserve the edits, then reopen the original DOCX as a separate document.');

function hydrateHtml(html: string, source: string) {
  const parser = new DOMParser(), current = parser.parseFromString(html, 'text/html');
  const before = parser.parseFromString(source, 'text/html');
  const originals = new Map([...before.querySelectorAll('[data-source-paragraph]')]
    .map(p => [p.getAttribute('data-source-paragraph'), p]));
  const textNodes = (root: Element) => {
    const nodes: Text[] = [], walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) nodes.push(walker.currentNode as Text);
    return nodes;
  };
  const semantic = (text: Text) => text.parentElement?.closest('[data-word-hyphen]') ? '\u00ad' : text.data;
  for (const paragraph of current.querySelectorAll('p,h1,h2,h3,h4,h5,h6')) {
    const nodes = textNodes(paragraph);
    if (!nodes.some(text => text.data.includes('\u00ad') && !text.parentElement?.closest('[data-word-hyphen]'))) continue;
    const original = originals.get(paragraph.getAttribute('data-source-paragraph'));
    if (!original) throw failure();
    let sourceText = '';
    const kinds = new Map<number, string>();
    for (const text of textNodes(original)) {
      const marker = text.parentElement?.closest('[data-word-hyphen]');
      if (marker) {
        const kind = marker.getAttribute('data-word-hyphen');
        if (wordHyphenText(kind) === null || text.data !== wordHyphenText(kind)) throw failure();
        kinds.set(sourceText.length, kind!);
      }
      sourceText += semantic(text);
    }
    const currentText = nodes.map(semantic).join('');
    let prefix = 0, suffix = 0;
    while (prefix < Math.min(sourceText.length, currentText.length) && sourceText[prefix] === currentText[prefix]) prefix++;
    while (suffix < sourceText.length - prefix && suffix < currentText.length - prefix &&
      sourceText[sourceText.length - suffix - 1] === currentText[currentText.length - suffix - 1]) suffix++;
    let offset = 0;
    for (const text of nodes) {
      if (!text.parentElement?.closest('[data-word-hyphen]') && text.data.includes('\u00ad')) {
        const fragment = current.createDocumentFragment();
        let start = 0;
        for (let i = 0; i < text.length; i++) if (text.data[i] === '\u00ad') {
          const position = offset + i;
          const at = position < prefix ? position
            : position >= currentText.length - suffix ? sourceText.length - currentText.length + position : -1;
          const kind = kinds.get(at);
          if (!kind) throw failure();
          fragment.append(current.createTextNode(text.data.slice(start, i)));
          const marker = current.createElement('span');marker.dataset.wordHyphen = kind;
          marker.textContent = wordHyphenText(kind)!;fragment.append(marker);start = i + 1;
        }
        fragment.append(current.createTextNode(text.data.slice(start)));text.replaceWith(fragment);
      }
      offset += semantic(text).length;
    }
  }
  // Unmapped legacy controls outside retained paragraphs cannot be guessed.
  if (textNodes(current.body).some(text => text.data.includes('\u00ad') && !text.parentElement?.closest('[data-word-hyphen]')))
    throw failure();
  return current.body.innerHTML;
}

/** Restore only controls in an unchanged prefix/suffix of the retained source.
 * The caller's content and original bytes stay untouched if any part fails. */
export function hydrateWordHyphens(current: WordContent, source: WordContent): WordContent {
  const html = hydrateHtml(current.html, source.html);
  const stories = current.stories ? { ...current.stories, parts: current.stories.parts.map(part => {
    const original = source.stories?.parts.find(p => p.path === (part.copiedFrom || part.path));
    let sourceHtml = original?.html;
    if (part.copiedFrom && original) sourceHtml = clonedStoryHtml(original, part.path);
    if (!part.html.includes('\u00ad')) return part;
    if (!sourceHtml) throw failure();
    return { ...part, html: hydrateHtml(part.html, sourceHtml) };
  }) } : current.stories;
  return { ...current, html, stories, hyphenVersion: 1 };
}
