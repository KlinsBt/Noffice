import type { WordLineSpacing } from './word-line-spacing';

/** Pinned Word's regular Inter baseline intervals across 12- and 80-line
 * native paragraphs, with the separately embedded default document included.
 * Columns: Single, 1.5, 1.65, Double, minimum18. Values are points, rounded to
 * half twips near the observed interval centers. Exact18 uses 14.45pt.
 * Native line advances remain fractional font units; centipoint rounding
 * failed the independent long-paragraph matrix. This is a bounded profile. */
const baselines = new Map([
  [8, [7.825, 7.85, 7.85, 7.75, 7.825]],
  [10, [9.75, 9.775, 9.65, 9.8, 9.75]],
  [12, [11.7, 11.675, 11.725, 11.7, 11.7]],
  [14, [13.625, 13.625, 13.625, 13.625, 13.65]],
  [18, [17.45, 17.45, 17.45, 17.6, 17.45]],
  [24, [23.35, 23.35, 23.35, 23.35, 23.35]],
]);

export function wordInterMetrics(
  sizePx: number,
  ratio: number,
  style: Pick<CSSStyleDeclaration, 'fontFamily' | 'fontWeight' | 'fontStyle' | 'fontStretch'>,
) {
  const family = style.fontFamily.split(',')[0].replace(/["']/g, '').trim().toLowerCase();
  const size = sizePx * 0.75;
  const measuredSize = [...baselines.keys()].find((n) => Math.abs(n - size) < 0.0001);
  if (
    family !== 'inter' ||
    !['normal', '400'].includes(style.fontWeight) ||
    style.fontStyle !== 'normal' ||
    !['normal', '100%'].includes(style.fontStretch) ||
    measuredSize === undefined ||
    !Number.isFinite(ratio) ||
    Math.abs(ratio - 2478 / 2048) > 0.0001
  )
    return null;
  const natural = (sizePx * 2478) / 2048;
  const supported = (spacing: WordLineSpacing) =>
    spacing.rule === 'auto' ? [240, 360, 396, 480].includes(spacing.line) : spacing.line === 360;
  return {
    supports: supported,
    natural,
    height: (spacing: WordLineSpacing) =>
      spacing.rule === 'exact'
        ? spacing.line / 15
        : spacing.rule === 'atLeast'
          ? Math.max(spacing.line / 15, natural)
          : (natural * spacing.line) / 240,
    ascent: (spacing: WordLineSpacing) =>
      spacing.rule === 'exact'
        ? 14.45 / 0.75
        : baselines.get(measuredSize)![
            spacing.rule === 'atLeast' ? 4 : [240, 360, 396, 480].indexOf(spacing.line)
          ] / 0.75,
  };
}
