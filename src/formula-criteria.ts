import type { Value } from './formulas';

export function wildcardMatch(pattern: string, value: Value): boolean {
  if (typeof value !== 'string') return false;
  let source = '';
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === '~' && i + 1 < pattern.length)
      source += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else source += c === '*' ? '.*' : c === '?' ? '.' : c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp('^' + source + '$', 'i').test(value);
}

/** Criteria are not ordinary formula comparisons: numeric text is coerced, wildcards
 * apply to text, and an error in a criteria range is a matchable value. */
export function criterionMatches(
  criterion: Value,
): (value: Value | Error, blank?: boolean) => boolean {
  const match = typeof criterion === 'string' ? /^(<=|>=|<>|=|<|>)(.*)$/.exec(criterion) : null;
  const op = match?.[1] || '=';
  const text = match ? match[2] : criterion;
  const rhs =
    typeof text === 'string' && text.trim() !== '' && Number.isFinite(Number(text))
      ? Number(text)
      : text;
  return (input, blank = false) => {
    if (input instanceof Error) {
      if (['#UNSUPPORTED!', '#LIMIT!', '#CYCLE!', '#ERROR!'].includes(input.message)) throw input;
      const equal = typeof rhs === 'string' && input.message.toLowerCase() === rhs.toLowerCase();
      return op === '=' ? equal : op === '<>' ? !equal : false;
    }
    if (op === '<>' && rhs === '') return !blank;
    const value =
      op === '=' &&
      typeof rhs === 'number' &&
      typeof input === 'string' &&
      input.trim() !== '' &&
      Number.isFinite(Number(input))
        ? Number(input)
        : input;
    const equal =
      typeof rhs === 'string'
        ? wildcardMatch(rhs, value)
        : typeof value === typeof rhs && value === rhs;
    if (op === '=') return equal;
    if (op === '<>') return !equal;
    if (typeof value !== typeof rhs) return false;
    const a = typeof value === 'string' ? value.toLowerCase() : value;
    const b = typeof rhs === 'string' ? rhs.toLowerCase() : rhs;
    return op === '<' ? a < b : op === '>' ? a > b : op === '<=' ? a <= b : a >= b;
  };
}
