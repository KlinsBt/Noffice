import type { WordTabStop } from './word-tab-stops';

/** CSS-pixel distance from the text margin; stop positions are semantic twips.
 * Automatic stops resume only beyond the last custom alignment stop. */
export function wordTabAdvance(
  current: number,
  following: number,
  decimalPrefix: number,
  stops: WordTabStop[],
  interval: number,
  indent: number,
  rightEdge = Infinity,
): { width: number; alignment: WordTabStop['alignment']; leader: WordTabStop['leader'] } | null {
  if (![current, following, decimalPrefix, interval, indent].every(Number.isFinite) || following < 0 || decimalPrefix < 0)
    return null;
  const custom = stops.filter((stop) => stop.alignment !== 'bar');
  if (custom.some((stop) => ['num', 'start', 'end'].includes(stop.alignment))) return null;
  const next = custom.find((stop) => stop.position / 15 > current + 1 / 64);
  // A hanging indent supplies an implicit stop before the first custom stop.
  if (indent > current + 1 / 64 && (!next || indent < next.position / 15))
    return { width: indent - current, alignment: 'left', leader: 'none' };
  if (next) {
    const aligned = next.alignment === 'right' ? following : next.alignment === 'center' ? following / 2 : next.alignment === 'decimal' ? decimalPrefix : 0;
    // Native center/right/decimal fields stop at the text margin when their
    // aligned end would cross it. The saved stop position remains unchanged.
    const start = next.alignment === 'left' ? next.position / 15
      : Math.min(next.position / 15 - aligned, rightEdge - following);
    const width = start - current;
    // Overfull aligned fields need the native wrapping/fallback contract.
    if (width < 0) return null;
    return { width, alignment: next.alignment, leader: next.leader };
  }
  if (interval <= 0) return null;
  const step = interval / 15;
  return { width: (Math.floor((current + 1 / 64) / step) + 1) * step - current, alignment: 'left', leader: 'none' };
}
