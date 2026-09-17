import type { WordContent } from './model';

/** Earlier imports read the preceding section when final properties were
 * absent. Explicit page choices and values distinct from that old baseline
 * remain user edits; only inherited control defaults are recovered. */
export function hydrateWordSectionDefaults(
  current: WordContent,
  source: WordContent,
  legacy: WordContent,
): WordContent {
  const content = { ...current, sectionDefaultsVersion: 1 as const };
  if (!current.pageOverrides?.paper && current.paper === legacy.paper)
    content.paper = source.paper;
  if (!current.pageOverrides?.margin && current.margin === legacy.margin)
    content.margin = source.margin;
  if (!current.pageOverrides?.orientation
    && (current.orientation || 'portrait') === (legacy.orientation || 'portrait'))
    content.orientation = source.orientation;
  return content;
}
