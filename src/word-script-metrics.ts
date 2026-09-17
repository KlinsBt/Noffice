import type { WordLineSpacing } from './word-line-spacing';

/** Independently captured Arial script profiles. Nominal document sizes remain
 * unchanged; only the derived view uses a smaller face and shifted baseline.
 * Other sizes, marks and line contexts need their own native qualification. */
export function wordScriptMetrics(
  kind: 'superscript' | 'subscript', size: number, paragraphSize: number,
  style: Pick<CSSStyleDeclaration, 'fontFamily' | 'fontWeight' | 'fontStyle' | 'fontStretch'>,
  spacing: WordLineSpacing,
  singleTextLine = false,
) {
  const points = size * .75;
  const context = (spacing.rule === 'exact' && spacing.line === 800) ||
    (singleTextLine && spacing.rule === 'auto' && spacing.line === 240 && Math.abs(points - 10) < .0001);
  if (Math.abs(size - paragraphSize) > .0001 || !context
    || style.fontFamily.split(',')[0].replace(/["']/g, '').trim().toLowerCase() !== 'arial'
    || !['normal', '400'].includes(style.fontWeight) || style.fontStyle !== 'normal'
    || !['normal', '100%'].includes(style.fontStretch)) return null;
  // Native PDFs resolve these nominal sizes to measured physical paint sizes.
  // An inferred 65% scale accumulates excess width in scripted words.
  const profile = Math.abs(points - 10) < .0001 ? { size: 6.48, shifts: [-3, .48] }
    : Math.abs(points - 20) < .0001 ? { size: 12.96, shifts: [-6.96, 1.56] } : null;
  return profile ? { kind, size: profile.size / .75,
    shift: profile.shifts[kind === 'superscript' ? 0 : 1] / .75 } : null;
}
