import { z } from 'zod';
import { wordSectionIdSchema } from './word-section-identity';
import type { WordSectionLayout } from './word-section-layout';

const twips = z.number().int().min(0).max(31680);
export const wordSectionGeometrySchema = z
  .object({
    sectionId: wordSectionIdSchema,
    width: twips.min(1),
    height: twips.min(1),
    orientation: z.enum(['portrait', 'landscape']),
    margins: z.object({ left: twips, right: twips, top: twips, bottom: twips }),
  })
  .refine(
    (value) =>
      value.margins.left + value.margins.right < value.width &&
      value.margins.top + value.margins.bottom < value.height,
    'Section margins must leave a positive text area.',
  );
export type WordSectionGeometry = z.infer<typeof wordSectionGeometrySchema>;
export const wordSectionGeometryChangeSchema = z.discriminatedUnion('key', [
  z.object({ key: z.literal('paper'), value: z.enum(['a4', 'letter']) }),
  z.object({ key: z.literal('margin'), value: z.enum(['normal', 'narrow', 'wide']) }),
  z.object({ key: z.literal('orientation'), value: z.enum(['portrait', 'landscape']) }),
]);
export type WordSectionGeometryChange = z.infer<typeof wordSectionGeometryChangeSchema>;

/** Freeze measured live geometry, never mutate retained source properties. */
export function wordSectionGeometry(section: WordSectionLayout): WordSectionGeometry {
  return wordSectionGeometrySchema.parse({
    sectionId: section.id,
    width: section.width,
    height: section.height,
    orientation: section.orientation,
    margins: Object.fromEntries(
      (['left', 'right', 'top', 'bottom'] as const).map((side) => [side, section.margins[side]]),
    ),
  });
}

export function changedWordSectionGeometry(
  current: WordSectionGeometry,
  input: WordSectionGeometryChange,
): WordSectionGeometry {
  const change = wordSectionGeometryChangeSchema.parse(input);
  const next = structuredClone(current);
  if (change.key === 'margin') {
    const value = change.value === 'narrow' ? 720 : change.value === 'wide' ? 2160 : 1440;
    next.margins = { left: value, right: value, top: value, bottom: value };
  } else {
    const dimensions =
      change.key === 'paper'
        ? change.value === 'letter'
          ? [12240, 15840]
          : // Installed Word's A4 PaperSize setter writes these integer twips;
            // an existing imported A4-like size is preserved until this command.
            [11907, 16839]
        : [next.width, next.height].sort((a, b) => a - b);
    if (change.key === 'orientation' && next.orientation !== change.value) {
      const { left, right, top, bottom } = next.margins;
      next.margins =
        change.value === 'landscape'
          ? { left: bottom, top: left, right: top, bottom: right }
          : { left: top, top: right, right: bottom, bottom: left };
      next.orientation = change.value;
    }
    // Native PageSetup.PaperSize uses the preset's portrait dimensions. Unlike
    // Orientation, that assignment does not rotate margins or section starts.
    if (change.key === 'paper') next.orientation = 'portrait';
    if (next.orientation === 'landscape') dimensions.reverse();
    [next.width, next.height] = dimensions;
  }
  return wordSectionGeometrySchema.parse(next);
}
