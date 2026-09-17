import type { WordStories } from './word-stories';
import type { WordSectionLayout } from './word-section-layout';

export type WordStoryMeasurements = Omit<WordStories, 'parts'> & {
  parts: (Omit<WordStories['parts'][number], 'html'> & {
    height: number;
    width?: number;
    pageLeft?: number;
    renderKey?: string;
  })[];
};
export interface WordStoryPlacement {
  path: string;
  kind: 'header' | 'footer';
  left: number;
  top: number;
  width: number;
  height: number;
  renderKey?: string;
}

/** Page-local CSS-pixel bounds. A long first/even story affects only pages on
 * which it is selected; Word does not reserve the largest story on every page. */
export function wordStoryPageBounds(
  section: WordSectionLayout,
  measurements: WordStoryMeasurements,
  pageNumber: number,
  firstInSection: boolean,
): { top: number; bottom: number; stories: WordStoryPlacement[] } | null {
  if (!Number.isSafeInteger(pageNumber) || pageNumber < 1) return null;
  const { width, height, margins } = section;
  if (
    width === null ||
    height === null ||
    [width, height, margins.top, margins.bottom, margins.left, margins.right].some(
      (n) => n === null || !Number.isFinite(n) || n < 0,
    )
  )
    return null;
  const slot =
    firstInSection && section.differentFirstPage
      ? 'first'
      : measurements.evenAndOddHeaders && pageNumber % 2 === 0
        ? 'even'
        : 'default';
  let top = margins.top! / 15,
    bottom = (height - margins.bottom!) / 15;
  const bodyWidth = (width - margins.left! - margins.right!) / 15;
  if (!(bodyWidth > 0)) return null;
  const stories: WordStoryPlacement[] = [];
  for (const kind of ['header', 'footer'] as const) {
    const reference = section[kind === 'header' ? 'headers' : 'footers'][slot];
    if (!reference) continue;
    const candidates = measurements.parts.filter(
      (part) =>
        part.relationshipIds.includes(reference.relationshipId) &&
        (part.width === undefined || part.width === bodyWidth) &&
        (part.pageLeft === undefined || part.pageLeft === margins.left! / 15),
    );
    const part = candidates[0],
      distance = margins[kind];
    if (
      candidates.length !== 1 ||
      part.kind !== kind ||
      !Number.isFinite(part.height) ||
      part.height <= 0 ||
      distance === null ||
      !Number.isFinite(distance) ||
      distance < 0
    )
      return null;
    const storyTop = kind === 'header' ? distance / 15 : (height - distance) / 15 - part.height;
    if (kind === 'header') top = Math.max(top, storyTop + part.height);
    else bottom = Math.min(bottom, storyTop);
    stories.push({
      path: part.path,
      kind,
      left: margins.left! / 15,
      top: storyTop,
      width: bodyWidth,
      height: part.height,
      ...(part.renderKey ? { renderKey: part.renderKey } : {}),
    });
  }
  return bottom > top ? { top, bottom, stories } : null;
}
