import SSF from 'ssf';
import { numericFormat } from './numeric-format';

export function formatNumber(code: string, value: string | number | boolean, date1904 = false) {
  if (code.length > 2048) throw new Error('#LIMIT!');
  if (!SSF.is_date(code)) {
    const result = numericFormat(code, value);
    if (result !== undefined) return result;
  }
  // Excel accepts bare date-separator dots; SSF mistakes them for fractional-second tokens.
  let normalized = code.replace(
    /"(?:[^"]|"")*"|\\.|([dmy])\.(?=[dmy])/gi,
    (match, dateToken: string | undefined) => (dateToken ? dateToken + '"."' : match),
  );
  let marker = 'NOFFICEFMT';
  while (code.includes(marker)) marker += 'X';
  const substitutions: { start: string; end: string; render: (s: string) => string }[] = [];
  normalized = normalized.replace(/"(?:[^"]|"")*"|\\.|a\/p|(?<!s)s(?=\.0{1,3}(?!0))/gi, (match) => {
    if (match.startsWith('"') || match.startsWith('\\')) return match;
    const start = marker + substitutions.length + 'BEGIN',
      end = marker + substitutions.length + 'END';
    substitutions.push({
      start,
      end,
      render: /a\/p/i.test(match)
        ? (text) => (text === 'A' ? match[0] : match[2])
        : (text) => text.replace(/^0(?=\d)/, ''),
    });
    return `"${start}"${/a\/p/i.test(match) ? 'A/P' : 'ss'}"${end}"`;
  });
  let result = SSF.format(normalized, value, { date1904 });
  for (const { start, end, render } of substitutions) {
    const from = result.indexOf(start),
      to = result.indexOf(end, from + start.length);
    if (from >= 0 && to >= 0)
      result =
        result.slice(0, from) +
        render(result.slice(from + start.length, to)) +
        result.slice(to + end.length);
  }
  return result;
}
