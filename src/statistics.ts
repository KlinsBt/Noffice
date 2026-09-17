/** Scalar statistical functions; reference arguments exclude text and logical cells. */
type Atom = string | number | boolean;
export type StatisticalArgument = { values: Atom[]; reference: boolean };
export const statisticalFunctions = new Set([
  'VAR',
  'VAR.S',
  'VARP',
  'VAR.P',
  'STDEV',
  'STDEV.S',
  'STDEVP',
  'STDEV.P',
  'MEDIAN',
  'LARGE',
  'SMALL',
  'RANK',
  'RANK.EQ',
  'RANK.AVG',
  'PERCENTILE',
  'PERCENTILE.INC',
  'PERCENTILE.EXC',
  'QUARTILE',
  'QUARTILE.INC',
  'QUARTILE.EXC',
  'PRODUCT',
  'GEOMEAN',
  'HARMEAN',
  'DEVSQ',
  'AVEDEV',
  'MODE',
  'MODE.SNGL',
]);
const fail = (error: string): never => {
  throw new Error(error);
};
function numeric(value: Atom): number {
  if (typeof value === 'boolean') return Number(value);
  if (typeof value === 'number') return value;
  if (!value.trim() || !Number.isFinite(Number(value))) return fail('#VALUE!');
  return Number(value);
}
export function statistic(name: string, args: StatisticalArgument[]): number {
  if (!args.length || args.length > 255) return fail('#VALUE!');
  const rank = name.startsWith('RANK'),
    ordered = ['LARGE', 'SMALL'].includes(name),
    percentile = name.startsWith('PERCENTILE'),
    quartile = name.startsWith('QUARTILE');
  if (
    (rank && (args.length < 2 || args.length > 3)) ||
    ((ordered || percentile || quartile) && args.length !== 2)
  )
    return fail('#VALUE!');
  if (rank && !args[1].reference) return fail('#VALUE!');
  const data = rank ? [args[1]] : ordered || percentile || quartile ? [args[0]] : args;
  if (
    name.startsWith('MODE') &&
    data.some((arg) => arg.values.length === 1 && typeof arg.values[0] !== 'number')
  )
    return fail('#VALUE!');
  const numbers = data.flatMap((arg) =>
    arg.values.flatMap((v) => (typeof v === 'number' ? [v] : arg.reference ? [] : [numeric(v)])),
  );
  const n = numbers.length;
  const param = (index: number) => numeric(args[index].values[0]);
  if (name === 'PRODUCT') return n ? numbers.reduce((a, b) => a * b, 1) : 0;
  if (!n)
    return fail(
      rank
        ? '#N/A'
        : ordered || percentile || quartile
          ? '#NUM!'
          : name.startsWith('MODE') || name === 'HARMEAN'
            ? '#N/A'
            : ['MEDIAN', 'GEOMEAN', 'AVEDEV', 'DEVSQ'].includes(name)
              ? '#NUM!'
              : '#DIV/0!',
    );
  if (rank) {
    const value = param(0),
      ascending = args[2] ? param(2) !== 0 : false;
    const count = numbers.filter((x) => x === value).length;
    if (!count) return fail('#N/A');
    return (
      1 +
      numbers.filter((x) => (ascending ? x < value : x > value)).length +
      (name === 'RANK.AVG' ? (count - 1) / 2 : 0)
    );
  }
  if (ordered || percentile || quartile || name === 'MEDIAN') {
    numbers.sort((a, b) => a - b);
    if (ordered) {
      const k = name === 'SMALL' ? Math.trunc(param(1)) : Math.ceil(param(1));
      if (k < 1 || k > n) return fail('#NUM!');
      return numbers[name === 'LARGE' ? n - k : k - 1];
    }
    let k = name === 'MEDIAN' ? 0.5 : param(1);
    if (quartile) {
      if (k < 0) return fail('#NUM!');
      k = Math.trunc(k) / 4;
    }
    const exclusive = name.endsWith('.EXC');
    const position = exclusive ? k * (n + 1) - 1 : k * (n - 1);
    if (k < 0 || k > 1 || (exclusive && (k <= 0 || k >= 1)) || position < 0 || position > n - 1)
      return fail('#NUM!');
    const lower = Math.floor(position),
      fraction = position - lower;
    return numbers[lower] + fraction * (numbers[Math.min(lower + 1, n - 1)] - numbers[lower]);
  }
  if (name.startsWith('MODE')) {
    const counts = new Map<number, number>();
    for (const value of numbers) counts.set(value, (counts.get(value) || 0) + 1);
    const best = Math.max(...counts.values());
    if (best < 2) return fail('#N/A');
    return [...counts].find(([, count]) => count === best)![0];
  }
  if (name === 'GEOMEAN' || name === 'HARMEAN') {
    if (numbers.some((x) => x <= 0)) return fail('#NUM!');
    return name === 'GEOMEAN'
      ? Math.exp(numbers.reduce((sum, x) => sum + Math.log(x), 0) / n)
      : n / numbers.reduce((sum, x) => sum + 1 / x, 0);
  }
  // Center on the first value before summing to reduce cancellation for large offsets.
  const origin = numbers[0],
    meanOffset = numbers.reduce((sum, x) => sum + (x - origin), 0) / n;
  const deviations = numbers.map((x) => x - origin - meanOffset);
  if (name === 'AVEDEV') return deviations.reduce((sum, x) => sum + Math.abs(x), 0) / n;
  const squares = deviations.reduce((sum, x) => sum + x * x, 0);
  if (name === 'DEVSQ') return squares;
  const sample = ['VAR', 'VAR.S', 'STDEV', 'STDEV.S'].includes(name);
  if (sample && n < 2) return fail('#DIV/0!');
  const variance = squares / (n - Number(sample));
  return name.startsWith('STDEV') ? Math.sqrt(variance) : variance;
}
