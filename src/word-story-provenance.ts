import type { WordStories } from './word-stories';

export const clonedStoryParagraphKey = (path: string, key: string) => `${path}#${key}`;

/** Reading IDs belong to a part, even when native unlink copies its whole XML. */
export function clonedStoryHtml(part: WordStories['parts'][number], path: string) {
  const doc = new DOMParser().parseFromString(part.html, 'text/html');
  for (const paragraph of doc.querySelectorAll('[data-source-paragraph]')) {
    let key = paragraph.getAttribute('data-source-paragraph')!;
    if (part.copiedFrom || part.created) {
      if (!key.startsWith(`${part.path}#`))
        throw Error('The copied header/footer has invalid paragraph provenance.');
      key = key.slice(part.path.length + 1);
    }
    paragraph.setAttribute('data-source-paragraph', clonedStoryParagraphKey(path, key));
  }
  return doc.body.innerHTML;
}
