export interface WordJustifiedLine {
  /** Semantic UTF-16 offsets, including the separator after a nonfinal line. */
  from: number;
  to: number;
  textEnd: number;
  natural: number;
  /** Additional advance per ordinary space; the last line is not expanded. */
  wordSpacing: number;
}

/** The measured modern Word line-choice policy for homogeneous Latin text.
 * Callers must qualify compatibility, font advances and formatting separately.
 * Caret coordinates and painted PDF positions are not logical font advances.
 * Tabs, repeated/edge spaces, long-word splitting and complex text are pending. */
export function wordJustifiedLines(
  text: string,
  width: number,
  measure: (text: string) => number,
): WordJustifiedLine[] | null {
  if (!text || text.length > 20000 || !/^[!-~]+(?: [!-~]+)*$/.test(text)
    || !Number.isFinite(width) || width <= 0) return null;
  const words = [...text.matchAll(/[^ ]+/g)].map((match) => ({
    from: match.index!, to: match.index! + match[0].length, width: measure(match[0]),
  }));
  const space = measure(' ');
  if (!Number.isFinite(space) || space <= 0 || words.some((word) =>
    !Number.isFinite(word.width) || word.width <= 0 || word.width > width)) return null;
  const lines: WordJustifiedLine[] = [];
  for (let start = 0; start < words.length;) {
    let end = start, natural = words[start].width;
    while (end + 1 < words.length && natural + space + words[end + 1].width <= width)
      natural += space + words[++end].width;
    if (end + 1 < words.length && end > start) {
      const next = natural + space + words[end + 1].width;
      const expansion = (width - natural) / ((end - start) * space);
      const compression = (width - next + (end + 1 - start) * space) / ((end + 1 - start) * space);
      if (compression >= .75 && expansion > 2 * (1 - compression)) {
        end++;
        natural = next;
      }
    }
    const last = end === words.length - 1;
    lines.push({
      from: words[start].from,
      to: last ? text.length : words[end + 1].from,
      textEnd: words[end].to,
      natural,
      wordSpacing: !last && end > start ? (width - natural) / (end - start) : 0,
    });
    start = end + 1;
  }
  return lines;
}
