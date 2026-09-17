/** The pinned 600dpi print path quantizes the paragraph advance before its
 * physical origin. Thirty-two independent margin/prefix cases distinguish
 * that two-stage rounding from rounding the final DOM coordinate alone.
 * Use only with the separately measured native leader font profiles. */
export function wordTabLeaderPrintBounds(pageLeft: number, advance: number, width: number): [number, number] | null {
  if (![pageLeft, advance, width].every(Number.isFinite) || width < 0) return null;
  const quantize = (points: number) => Math.round(points / wordTabLeaderPrinterStep) * wordTabLeaderPrinterStep;
  const start = quantize(pageLeft + quantize(advance));
  const end = quantize(pageLeft + advance + width);
  return Number.isFinite(start) && Number.isFinite(end) && end >= start ? [start, end] : null;
}

/** Physical page points. Native leader glyphs occupy complete printer-advance
 * cells on the page grid, excluding the partial cells at either tab endpoint.
 * The caller supplies independently resolved font advance and physical bounds;
 * this function does not infer fonts, wrapping or story/page origins. */
export function wordTabLeaderOrigins(start: number, end: number, pitch: number): number[] | null {
  if (![start, end, pitch].every(Number.isFinite) || pitch <= 0 || end < start) return null;
  // Avoid moving an exactly integral quotient across its boundary through
  // floating-point division. This allowance scales only with machine precision.
  const quotient = (value: number) => {
    const q = value / pitch, nearest = Math.round(q);
    return Math.abs(q - nearest) <= Number.EPSILON * Math.max(1, Math.abs(q)) * 4 ? nearest : q;
  };
  const first = Math.ceil(quotient(start)), last = Math.floor(quotient(end));
  if (!Number.isSafeInteger(first) || !Number.isSafeInteger(last)) return null;
  const count = Math.max(0, last - first);
  if (count > 5000) return null;
  return Array.from({ length: count }, (_, i) => (first + i) * pitch);
}
export const wordTabLeaderPrinterStep = 72 / 600;
