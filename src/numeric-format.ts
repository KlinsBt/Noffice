type Token = { kind: 'code' | 'literal' | 'meta'; text: string };
const placeholder = (t: Token) => t.kind === 'code' && /^[0#?]$/.test(t.text);

function tokenize(format: string): Token[][] {
  const sections: Token[][] = [[]];
  for (let i = 0; i < format.length; i++) {
    const ch = format[i],
      tokens = sections.at(-1)!;
    if (ch === ';') {
      sections.push([]);
      continue;
    }
    if (ch === '"') {
      let text = '';
      while (++i < format.length && format[i] !== '"') text += format[i];
      if (i === format.length) throw new Error('#VALUE!');
      tokens.push({ kind: 'literal', text });
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      if (++i === format.length) throw new Error('#VALUE!');
      tokens.push({ kind: 'literal', text: ch === '\\' ? format[i] : ch === '_' ? ' ' : '' });
    } else if (ch === '[') {
      const end = format.indexOf(']', i + 1);
      if (end < 0) throw new Error('#VALUE!');
      tokens.push({ kind: 'meta', text: format.slice(i + 1, end) });
      i = end;
    } else tokens.push({ kind: 'code', text: ch });
  }
  if (sections.length > 4) throw new Error('#VALUE!');
  return sections;
}

/** Decimal shifting uses Excel's 15 significant decimal digits, avoiding binary
 * half-way artifacts when percentages and comma scaling are applied. */
function rounded(value: number, shift: number, decimals: number): string {
  const [mantissa, exp = '0'] = Math.abs(value).toPrecision(15).split('e');
  const digits = mantissa.replace('.', '');
  const power = Number(exp) - (mantissa.split('.')[1] || '').length + shift + decimals;
  let integer = BigInt(digits);
  if (power >= 0) integer *= 10n ** BigInt(power);
  else {
    const divisor = 10n ** BigInt(-power);
    integer = (integer + divisor / 2n) / divisor;
  }
  return integer.toString().padStart(decimals + 1, '0');
}

function integerText(tokens: Token[], value: string, suppressZero = true): string {
  const slots = tokens.filter(placeholder);
  let digits = suppressZero && /^0+$/.test(value) ? '' : value;
  let index = digits.length - 1;
  const grouped = tokens.some(
    (t, i) =>
      t.kind === 'code' &&
      t.text === ',' &&
      tokens.slice(0, i).some(placeholder) &&
      tokens.slice(i + 1).some(placeholder),
  );
  if (grouped && tokens.every((t) => placeholder(t) || (t.kind === 'code' && t.text === ','))) {
    const width = Math.max(digits.length, slots.length);
    let text = '';
    for (let n = 0; n < width; n++) {
      const slot = slots[slots.length - 1 - n];
      const digit =
        index >= 0 ? digits[index--] : slot?.text === '0' ? '0' : slot?.text === '?' ? ' ' : '';
      if (n && n % 3 === 0 && digit) text = (digit === ' ' ? ' ' : ',') + text;
      text = digit + text;
    }
    return text;
  }
  const output = tokens.map((t) => t.text);
  for (let i = tokens.length - 1; i >= 0; i--) {
    if (!placeholder(tokens[i])) continue;
    output[i] =
      index >= 0
        ? digits[index--]
        : tokens[i].text === '0'
          ? '0'
          : tokens[i].text === '?'
            ? ' '
            : '';
  }
  const first = tokens.findIndex(placeholder);
  if (index >= 0 && first >= 0) output[first] = digits.slice(0, index + 1) + output[first];
  return output.join('');
}

function fractionText(tokens: Token[], value: number, slash: number): string | undefined {
  const isDigit = (t: Token) => t.kind === 'code' && /^[0-9#?]$/.test(t.text);
  const numeratorEnd = tokens.findLastIndex((t, i) => i < slash && placeholder(t));
  let numeratorStart = numeratorEnd;
  while (numeratorStart > 0 && placeholder(tokens[numeratorStart - 1])) numeratorStart--;
  const denominatorStart = tokens.findIndex((t, i) => i > slash && isDigit(t));
  if (numeratorEnd < 0 || denominatorStart < 0) return undefined;
  let denominatorEnd = denominatorStart;
  while (denominatorEnd + 1 < tokens.length && isDigit(tokens[denominatorEnd + 1]))
    denominatorEnd++;
  if (tokens.slice(denominatorEnd + 1).some(placeholder)) return undefined;
  const numerator = tokens.slice(numeratorStart, numeratorEnd + 1);
  const denominator = tokens.slice(denominatorStart, denominatorEnd + 1);
  const integerEnd = tokens.findLastIndex((t, i) => i < numeratorStart && placeholder(t));
  const integerStart = tokens.findIndex((t, i) => i <= integerEnd && placeholder(t));
  const mixed = integerEnd >= 0;
  const integer = mixed ? tokens.slice(integerStart, integerEnd + 1) : [];
  const separator = mixed ? tokens.slice(integerEnd + 1, numeratorStart) : [];
  const before = tokens
    .slice(0, mixed ? integerStart : numeratorStart)
    .map((t) => t.text)
    .join('');
  const divider = tokens
    .slice(numeratorEnd + 1, denominatorStart)
    .map((t) => t.text)
    .join('');
  const after = tokens
    .slice(denominatorEnd + 1)
    .map((t) => t.text)
    .join('');
  const fixed = denominator.some((t) => /[1-9]/.test(t.text));
  const max = fixed
    ? Number(denominator.map((t) => t.text).join(''))
    : 10 ** denominator.length - 1;
  if (!Number.isFinite(max) || max < 1 || max > (fixed ? 1000000 : 9999))
    throw new Error('#LIMIT!');
  let whole = mixed ? Math.floor(Math.abs(value)) : 0;
  const part = Math.abs(value) - whole;
  let num = 0,
    den = 1,
    error = Infinity;
  for (let d = fixed ? max : 1; d <= max; d++) {
    const n = Math.round(part * d),
      difference = Math.abs(part - n / d);
    if (difference < error - 1e-14) {
      num = n;
      den = d;
      error = difference;
    }
    if (error < 1e-14 || fixed) break;
  }
  if (mixed && num >= den) {
    whole += Math.floor(num / den);
    num %= den;
  }
  const displayFraction = !mixed || num !== 0 || numerator.some((t) => t.text === '0');
  const sep = separator.map((t) => t.text).join('');
  const integerZero =
    (value === 0 && integer.some((t) => t.text === '?')) ||
    !displayFraction ||
    integer.some((t) => t.text === '0');
  const head = mixed ? integerText(integer, String(whole), !integerZero) : '';
  if (!displayFraction) {
    const reserve = [...integer, ...numerator, ...denominator].some((t) => t.text === '?');
    return (
      before +
      head +
      (reserve
        ? ' '.repeat(sep.length + numerator.length + divider.length + denominator.length)
        : '') +
      after
    );
  }
  const join =
    whole !== 0 || integerZero
      ? sep
      : [...integer, ...numerator].some((t) => t.text === '?')
        ? ' '.repeat(sep.length)
        : '';
  const paddedDen = fixed ? String(den) : integerText(denominator, String(den), false);
  const spaces = /^ */.exec(paddedDen)![0];
  const denText = paddedDen.slice(spaces.length) + spaces;
  return (
    before + head + join + integerText(numerator, String(num), false) + divider + denText + after
  );
}

/** Returns undefined for format families delegated to SSF (dates/scientific/etc.). */
export function numericFormat(code: string, value: string | number | boolean): string | undefined {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  const sections = tokenize(code);
  if (typeof value === 'string') {
    const text = sections[3];
    if (!text) return value;
    return text
      .filter((t) => t.kind !== 'meta')
      .map((t) => (t.kind === 'code' && t.text === '@' ? value : t.text))
      .join('');
  }
  if (!Number.isFinite(value)) throw new Error('#NUM!');
  const condition = (tokens: Token[]) =>
    tokens.find((t) => t.kind === 'meta' && /^(<=|>=|<>|=|<|>)/.test(t.text))?.text;
  const hasCondition = sections.slice(0, 2).some(condition);
  let index = 0;
  if (hasCondition) {
    index = sections.slice(0, 2).findIndex((tokens) => {
      const clause = condition(tokens);
      if (!clause) return false;
      const m = /^(<=|>=|<>|=|<|>)(-?\d+(?:\.\d*)?(?:E[+-]?\d+)?)$/i.exec(clause);
      if (!m) return false;
      const n = Number(m[2]);
      return m[1] === '<'
        ? value < n
        : m[1] === '>'
          ? value > n
          : m[1] === '<='
            ? value <= n
            : m[1] === '>='
              ? value >= n
              : m[1] === '='
                ? value === n
                : value !== n;
    });
    if (index < 0)
      index =
        condition(sections[0]) && sections[1] && condition(sections[1])
          ? 2
          : condition(sections[0])
            ? 1
            : 0;
    if (!sections[index]) return undefined;
  } else index = value < 0 && sections.length > 1 ? 1 : value === 0 && sections.length > 2 ? 2 : 0;
  let tokens = sections[index];
  if (
    tokens.some(
      (t) =>
        t.kind === 'meta' &&
        !/^(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\s*\d+|[<>=].*|\$.*)$/i.test(
          t.text,
        ),
    )
  )
    return undefined;
  tokens = tokens.flatMap((t): Token[] =>
    t.kind !== 'meta'
      ? [t]
      : t.text.startsWith('$')
        ? [{ kind: 'literal', text: t.text.slice(1).replace(/-[\da-f]+$/i, '') }]
        : [],
  );
  if (tokens.some((t) => t.kind === 'code' && /[a-zA-Z@]/.test(t.text))) return undefined;
  if (!tokens.some(placeholder)) return tokens.map((t) => t.text).join('');
  const selectedCondition = condition(sections[index]);
  const negativeOnly =
    selectedCondition && /^(<=|<|=)(-?\d+(?:\.\d*)?(?:E[+-]?\d+)?)$/i.exec(selectedCondition);
  const suppressSign = negativeOnly && Number(negativeOnly[2]) <= 0;
  const negative = value < 0 && (hasCondition ? !suppressSign : sections.length === 1);
  const prefix = negative ? '-' : '';
  const slash = tokens.findIndex((t) => t.kind === 'code' && t.text === '/');
  if (slash >= 0) {
    const result = fractionText(tokens, value, slash);
    return result === undefined ? undefined : prefix + result;
  }
  const dots = tokens.filter((t) => t.kind === 'code' && t.text === '.');
  if (dots.length > 1) return undefined;
  const dot = tokens.findIndex((t) => t.kind === 'code' && t.text === '.');
  const split = dot < 0 ? tokens.length : dot;
  const percent = tokens.filter((t) => t.kind === 'code' && t.text === '%').length;
  let commas = 0;
  tokens = tokens.filter((t, i) => {
    if (t.kind !== 'code' || t.text !== ',') return true;
    const segmentEnd = i < split ? split : tokens.length;
    if (tokens.slice(i + 1, segmentEnd).some(placeholder)) return true;
    if (!tokens.slice(0, i).some(placeholder)) return true;
    commas++;
    return false;
  });
  const decimalAt = tokens.findIndex((t) => t.kind === 'code' && t.text === '.');
  const whole = decimalAt < 0 ? tokens : tokens.slice(0, decimalAt);
  const decimal = decimalAt < 0 ? [] : tokens.slice(decimalAt + 1);
  const places = decimal.filter(placeholder).length;
  if (places > 100 || percent > 100 || commas > 100) throw new Error('#LIMIT!');
  const digits = rounded(value, percent * 2 - commas * 3, places);
  const integer = places ? digits.slice(0, -places) : digits;
  const fraction = places ? digits.slice(-places) : '';
  let last = fraction.length - 1;
  while (last >= 0 && fraction[last] === '0') last--;
  let n = 0;
  const tail = decimal
    .map((t) => {
      if (!placeholder(t)) return t.text;
      const i = n++;
      return i <= last || t.text === '0' ? fraction[i] : t.text === '?' ? ' ' : '';
    })
    .join('');
  // Separate literal prefix/suffix so grouping is independent of currency labels.
  const first = whole.findIndex(placeholder),
    end = whole.findLastIndex(placeholder);
  const head =
    first < 0
      ? whole.map((t) => t.text).join('') + (integer === '0' ? '' : integer)
      : whole
          .slice(0, first)
          .map((t) => t.text)
          .join('') +
        integerText(whole.slice(first, end + 1), integer) +
        whole
          .slice(end + 1)
          .map((t) => t.text)
          .join('');
  return (/^0+$/.test(digits) ? '' : prefix) + head + (decimalAt < 0 ? '' : '.') + tail;
}
