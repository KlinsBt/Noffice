import { wildcardTextMatches } from './search-wildcards';
export type SearchOptions = {
  matchCase?: boolean;
  wholeWord?: boolean;
  wholeCell?: boolean;
  wildcards?: boolean;
  forReplacement?: boolean;
};
export type TextMatch = { from: number; to: number };
export type SearchBudget = { characters: number; matches: number; operations?: number };
const wordCharacter = /[\p{L}\p{N}\p{M}_]/u;

function before(text: string, index: number) {
  const last = text.charCodeAt(index - 1);
  return text.slice(index - (last >= 0xdc00 && last <= 0xdfff ? 2 : 1), index);
}

/** Literal search: regex syntax and replacement dollar sequences are ordinary text. */
export function textMatches(
  text: string,
  query: string,
  options: SearchOptions = {},
  budget: SearchBudget = { characters: 0, matches: 0 },
): TextMatch[] {
  if (!query) return [];
  if (query.length > 256) throw Error('Search text is limited to 256 characters.');
  budget.characters += text.length;
  if (budget.characters > 8000000) throw Error('This search is too large. Narrow its scope.');
  if (options.wildcards && /[?*~]/.test(query))
    return wildcardTextMatches(
      text,
      query,
      !!options.matchCase,
      !!options.wholeCell,
      budget,
      !!options.forReplacement,
    );
  const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(escaped, options.matchCase ? 'gu' : 'giu');
  const found: TextMatch[] = [];
  for (const match of text.matchAll(regex)) {
    const from = match.index!,
      to = from + match[0].length;
    if (options.wholeCell && (from !== 0 || to !== text.length)) continue;
    if (
      options.wholeWord &&
      ((from > 0 && wordCharacter.test(before(text, from))) ||
        (to < text.length && wordCharacter.test(String.fromCodePoint(text.codePointAt(to)!))))
    )
      continue;
    if (++budget.matches > 10000) throw Error('More than 10,000 matches. Narrow the search.');
    found.push({ from, to });
  }
  return found;
}

export function replaceMatches(
  text: string,
  matches: TextMatch[],
  replacement: string,
  maxLength: number,
) {
  if (
    text.length +
      matches.reduce((sum, match) => sum + replacement.length - (match.to - match.from), 0) >
    maxLength
  )
    throw Error(
      `Replacement would exceed the ${maxLength.toLocaleString('en-US')}-character limit.`,
    );
  let cursor = 0;
  const pieces: string[] = [];
  for (const match of matches) {
    if (match.from < cursor || match.to < match.from || match.to > text.length)
      throw Error('Search results are no longer valid.');
    pieces.push(text.slice(cursor, match.from), replacement);
    cursor = match.to;
  }
  pieces.push(text.slice(cursor));
  return pieces.join('');
}
