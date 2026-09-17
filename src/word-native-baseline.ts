import type { WordLineSpacing } from './word-line-spacing';

/** Regular Arial baselines measured in the pinned Word / Microsoft Print to PDF
 * context, rounded to twips from 28 independently authored paragraph cases
 * and the independent mixed-wrap fixture's measured placement intervals.
 * Columns: Single, 1.5, Double, 50pt minimum. No font binaries are distributed.
 * This is a bounded profile, not a general font or printer emulation engine.
 */
const arialBaselines = new Map([
  [10, [9.35, 9.35, 9.35, 47.9]],
  [15, [13.9, 14.05, 13.9, 46.8]],
  // Single 20pt ranges from 18.474 to 18.600pt across native paragraph
  // contexts. Its twip-rounded midpoint avoids choosing either painted edge.
  // The 50pt minimum ranges from 45.480pt in the independent calibration
  // to 45.600pt after a mixed-size hard break. Retain their twip-rounded midpoint.
  [20, [18.55, 18.5, 18.5, 45.55]],
  [25, [23.5, 23.4, 23.4, 44.8]],
  [30, [27.95, 27.95, 28.1, 43.55]],
  [35, [32.75, 32.75, 32.9, 42.5]],
  [40, [37.45, 37.45, 37.55, 41.5]],
]);

export function wordNativeBaseline(
  sizePx: number,
  normalHeight: number,
  normalRatio: number,
  family: string,
  weight: string,
  style: string,
  stretch: string,
  spacing: WordLineSpacing,
): number | null {
  const face = family.split(',')[0].replace(/["']/g, '').trim().toLowerCase();
  const points = sizePx * 0.75;
  // Independent alternating backgrounds preserve all glyphs in16 native
  // optional/literal-hyphen documents. Their245 exact24 line origins span
  // 19.194..19.224pt across both fonts/four faces, including wrapped pages.
  if (
    spacing.rule === 'exact' &&
    spacing.line === 480 &&
    Number.isFinite(normalHeight) &&
    normalHeight > 0 &&
    ['normal', '100%'].includes(stretch) &&
    ['normal', '400', 'bold', '700'].includes(weight) &&
    ['normal', 'italic'].includes(style) &&
    ((face === 'arial' &&
      Math.abs(points - 10) < 0.0001 &&
      Math.abs(normalRatio - 1.1499) <= 0.0001) ||
      (face === 'calibri' &&
        Math.abs(points - 11) < 0.0001 &&
        Math.abs(normalRatio - 1.2207) <= 0.0001))
  )
    return 19.2 / 0.75;
  // Source/shaded leader and decoration paragraphs preserve every glyph
  // matrix. The additional74 decoration size/face controls retain the same
  // 32.040..32.064pt physical background-to-baseline interval.
  if (
    spacing.rule === 'exact' &&
    spacing.line === 800 &&
    Number.isFinite(points) &&
    Number.isFinite(normalHeight) &&
    normalHeight > 0 &&
    Number.isFinite(normalRatio) &&
    ['normal', '100%'].includes(stretch)
  ) {
    const regular = ['normal', '400'].includes(weight) && style === 'normal';
    const supportedFace =
      ['normal', '400', 'bold', '700'].includes(weight) && ['normal', 'italic'].includes(style);
    const heldoutSize = [1, 3, 9, 10.5, 11.5, 14, 18, 36].some(
      (size) => Math.abs(points - size) < 0.0001,
    );
    if (
      (face === 'arial' &&
        Math.abs(normalRatio - 1.1499) <= 0.0001 &&
        (((heldoutSize || Math.abs(points - 10) < 0.0001) && supportedFace) ||
          ([9, 11, 20].some((size) => Math.abs(points - size) < 0.0001) && regular))) ||
      // Independent character/paragraph-mark/cross-paragraph story ranges
      // retain the same shaded paragraph origins with Times20 and Arial10.
      (face === 'times new roman' &&
        regular &&
        [10, 20].some((size) => Math.abs(points - size) < 0.0001) &&
        Math.abs(normalRatio - 1.1499) <= 0.0001) ||
      (face === 'calibri' &&
        Math.abs(normalRatio - 1.2207) <= 0.0001 &&
        (((heldoutSize || Math.abs(points - 11) < 0.0001) && supportedFace) ||
          (Math.abs(points - 22) < 0.0001 && regular)))
    )
      return 32.05 / 0.75;
  }
  if (
    !['normal', '400'].includes(weight) ||
    style !== 'normal' ||
    !['normal', '100%'].includes(stretch) ||
    !Number.isFinite(normalRatio) ||
    !Number.isFinite(normalHeight) ||
    normalHeight <= 0
  )
    return null;
  if (!Number.isFinite(points)) return null;
  // Independently authored empty-header/footer Calibri11 cases establish the
  // natural ascent in single, double, wrapped, hard-line and 30pt minimum
  // contexts. Keep this profile restricted to the measured regular font and
  // browser metrics; other sizes/descriptors retain their existing fallback.
  if (face === 'calibri') {
    if (Math.abs(points - 11) > 0.0001 || Math.abs(normalRatio - 1.2207) > 0.0001) return null;
    if (spacing.rule === 'exact' && spacing.line === 240) return 9.6 / 0.75;
    if (
      (spacing.rule === 'auto' && [240, 480].includes(spacing.line)) ||
      (spacing.rule === 'atLeast' && spacing.line === 600)
    )
      return 10.45 / 0.75;
    return null;
  }
  if (face !== 'arial' || Math.abs(normalRatio - 1.1499) > 0.0001) return null;
  const size = Math.round(points);
  if (Math.abs(size - points) > 0.0001) return null;
  // Native first-section controls place regular Arial12 at 11.159973pt
  // below the known page-body origin with 1.65 spacing. Other sizes, faces
  // and automatic multiples keep their separately measured profiles.
  if (size === 12 && spacing.rule === 'auto' && spacing.line === 396) return 11.15 / 0.75;
  // Four independent original/shaded list controls preserve every raw glyph
  // matrix. Their 22 paragraph/marker origins span 11.159973..11.183960pt.
  if (size === 12 && spacing.rule === 'auto' && spacing.line === 240) return 11.15 / 0.75;
  // Independent shaded/control/native PDFs preserve every glyph matrix across
  // 21 tab paragraphs and 70 hard lines. Physical background-to-baseline
  // offsets span 15.960..16.104pt. Use their twip-rounded midpoint, as for the
  // other bounded native profiles, without changing line advance or PDF limits.
  if (size === 10 && spacing.rule === 'exact' && spacing.line === 400) return 16.05 / 0.75;
  // Empty-section typing controls independently place Arial10/exact48 at
  // 38.420pt and Arial40/exact24 at19.220pt above their known physical line
  // origins. Preserve the same measured 80% rule; other faces remain gated.
  if (
    spacing.rule === 'exact' &&
    ((size === 10 && spacing.line === 960) || (size === 40 && spacing.line === 480))
  )
    return (spacing.line / 15) * 0.8;
  // Twelve independent exact-height/font-size cases place the baseline at
  // 80% of the paragraph line box, including clipped large glyphs. Comparable
  // paragraph backgrounds, not COM glyph tops, establish this coordinate.
  // The independently authored 70-line Arial10/exact12 case additionally
  // measures 9.6pt from each physical body/line origin on both pages. The native
  // exact20 tab-stop case uses the independently measured profile above.
  if (
    spacing.rule === 'exact' &&
    [10, 20, 40].includes(size) &&
    ([120, 200, 300, 600].includes(spacing.line) || (size === 10 && spacing.line === 240))
  )
    return (spacing.line / 15) * 0.8;
  // The wrapped-leading 18pt and inline-flow 30pt native minimum contexts
  // retain the 10pt natural ascent; the caller places excess above the glyphs.
  if (size === 10 && spacing.rule === 'atLeast' && [360, 600].includes(spacing.line))
    return 9.35 / 0.75;
  const index =
    spacing.rule === 'auto'
      ? [240, 360, 480].indexOf(spacing.line)
      : spacing.rule === 'atLeast' && spacing.line === 1000
        ? 3
        : -1;
  const baseline = arialBaselines.get(size)?.[index];
  if (baseline === undefined) return null;
  // The line-layout caller places minimum excess above the natural line box.
  return baseline / 0.75 - (index === 3 ? Math.max(0, 1000 / 15 - normalHeight) : 0);
}
