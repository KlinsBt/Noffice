import type { SearchBudget, TextMatch } from './text-search';

/** Excel consumes the shortest internal star span, but a trailing star consumes the tail. */
export function wildcardTextMatches(
  text: string,
  query: string,
  matchCase: boolean,
  wholeCell: boolean,
  budget: SearchBudget,
  forReplacement = false,
): TextMatch[] {
  const pattern = Array.from(query),
    characters = Array.from(text);
  const tokens: { type: 'literal' | 'star' | 'any'; text?: string }[] = [];
  for (let i = 0; i < pattern.length; i++) {
    const char = pattern[i];
    if (forReplacement && char === '~' && i + 1 === pattern.length) return [];
    if (char === '~' && i + 1 < pattern.length)
      tokens.push({ type: 'literal', text: pattern[++i] });
    else if (char === '*') {
      if (tokens.at(-1)?.type !== 'star') tokens.push({ type: 'star' });
    } else tokens.push(char === '?' ? { type: 'any' } : { type: 'literal', text: char });
  }
  budget.operations = (budget.operations || 0) + (characters.length + 1) * tokens.length;
  if (budget.operations > 4000000)
    throw Error('Wildcard search is too complex. Narrow its scope or use fewer wildcards.');
  let next = Int32Array.from({ length: characters.length + 1 }, (_, i) =>
    wholeCell && i !== characters.length ? -1 : i,
  );
  for (let i = tokens.length - 1; i >= 0; i--) {
    const current = new Int32Array(characters.length + 1).fill(-1),
      token = tokens[i];
    for (let j = characters.length; j >= 0; j--) {
      if (token.type === 'star')
        current[j] =
          i === tokens.length - 1
            ? characters.length
            : next[j] >= 0
              ? next[j]
              : j < characters.length
                ? current[j + 1]
                : -1;
      else if (
        j < characters.length &&
        (token.type === 'any' ||
          (matchCase
            ? token.text === characters[j]
            : token.text!.toLowerCase() === characters[j].toLowerCase() ||
              token.text!.toUpperCase() === characters[j].toUpperCase()))
      )
        current[j] = next[j + 1];
    }
    next = current;
  }
  const offsets = [0];
  for (const character of characters) offsets.push(offsets.at(-1)! + character.length);
  const found: TextMatch[] = [];
  for (let start = 0; start < characters.length;) {
    const end = next[start];
    if (end > start && (!wholeCell || (start === 0 && end === characters.length))) {
      if (++budget.matches > 10000) throw Error('More than 10,000 matches. Narrow the search.');
      found.push({ from: offsets[start], to: offsets[end] });
      start = end;
    } else start++;
  }
  return found;
}
