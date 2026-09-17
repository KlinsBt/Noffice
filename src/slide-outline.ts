import type { SlideElement } from './model';

export const outlineDashes = [
  'solid',
  'dot',
  'dash',
  'lgDash',
  'dashDot',
  'lgDashDot',
  'lgDashDotDot',
  'sysDash',
  'sysDot',
  'sysDashDot',
  'sysDashDotDot',
] as const;
export type SlideOutline = NonNullable<SlideElement['outline']>;
const patterns: Record<SlideOutline['dash'], number[]> = {
  solid: [],
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};
export function outlineDashArray(outline: SlideOutline) {
  return patterns[outline.dash].map((n) => n * outline.width).join(' ') || undefined;
}
export function editOutline(element: SlideElement, patch: Partial<SlideOutline>): SlideOutline {
  return { color: '#000000', width: 1, opacity: 1, dash: 'solid', ...element.outline, ...patch };
}
