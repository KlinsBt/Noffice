export type WordLineRule = 'auto' | 'exact' | 'atLeast';
export interface WordLineSpacing {
  rule: WordLineRule;
  line: number;
}

/** Native numeric entry uses Single-precision points, then rounds to twips.
 * Its editable range is 14..31680; retained OOXML below that range remains
 * readable through wordLineSpacing without becoming a valid new edit. */
export function wordLineSpacingInput(amount: unknown, rule: WordLineRule): WordLineSpacing | null {
  if (typeof amount !== 'number' || !Number.isFinite(amount) || amount <= 0 ||
      !['auto', 'exact', 'atLeast'].includes(rule)) return null;
  const points = rule === 'auto' ? Math.fround(Math.fround(amount) * 12) : Math.fround(amount);
  const line = Math.round(points * 20);
  return Number.isSafeInteger(line) && line >= 14 && line <= 31680 ? { rule, line } : null;
}

/** Existing physical HTML values mean exact spacing. Only an explicit minimum
 * marker changes that meaning; unitless values always mean automatic multiples. */
export function wordLineSpacing(value: unknown, minimum?: unknown): WordLineSpacing | null {
  if (typeof value !== 'string') return null;
  const match = /^(\d+(?:\.\d+)?)(pt|px)?$/.exec(value);
  if (!match) return null;
  const line = Math.round(
    Number(match[1]) * (match[2] === 'pt' ? 20 : match[2] === 'px' ? 15 : 240),
  );
  if (!Number.isFinite(line) || line < 1 || line > 31680) return null;
  return { rule: match[2] ? (minimum === 'atLeast' ? 'atLeast' : 'exact') : 'auto', line };
}
