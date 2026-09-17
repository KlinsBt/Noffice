/** Constant-period financial functions. Inputs are coerced by the formula evaluator. */
export const financialFunctions = new Set([
  'PMT',
  'PV',
  'FV',
  'NPER',
  'IPMT',
  'PPMT',
  'SLN',
  'SYD',
  'EFFECT',
  'NOMINAL',
]);
const fail = (code = '#NUM!'): never => {
  throw new Error(code);
};
function growth(rate: number, periods: number) {
  // log1p/expm1 avoid cancellation for very small rates; negative bases need pow.
  const power = rate > -1 ? Math.exp(periods * Math.log1p(rate)) : (1 + rate) ** periods;
  const annuity =
    rate === 0
      ? periods
      : rate > -1
        ? Math.expm1(periods * Math.log1p(rate)) / rate
        : (power - 1) / rate;
  if (!Number.isFinite(power) || !Number.isFinite(annuity)) fail();
  return { power, annuity };
}
function payment(rate: number, periods: number, present: number, future: number, due: number) {
  if (rate <= -1 || periods === 0) return fail();
  const { power, annuity } = growth(rate, periods);
  const divisor = (1 + rate * due) * annuity;
  return divisor ? -(present * power + future) / divisor : fail();
}
function futureValue(rate: number, periods: number, pmt: number, present: number, due: number) {
  const { power, annuity } = growth(rate, periods);
  return -(present * power + pmt * (1 + rate * due) * annuity);
}
export function financial(name: string, values: number[]): number {
  const [a, b, c, d = 0, e = 0, f = 0] = values;
  const count = values.length;
  const bounds =
    name === 'SLN'
      ? [3, 3]
      : name === 'SYD'
        ? [4, 4]
        : ['EFFECT', 'NOMINAL'].includes(name)
          ? [2, 2]
          : ['IPMT', 'PPMT'].includes(name)
            ? [4, 6]
            : [3, 5];
  if (count < bounds[0] || count > bounds[1]) return fail('#VALUE!');
  if (name === 'SLN') return c === 0 ? fail('#DIV/0!') : (a - b) / c;
  if (name === 'SYD') {
    if (c <= 0 || d <= 0 || d > c) return fail();
    return ((a - b) * (c - d + 1) * 2) / (c * (c + 1));
  }
  if (name === 'EFFECT' || name === 'NOMINAL') {
    const periods = Math.trunc(b);
    if (a <= 0 || periods < 1) return fail();
    return name === 'EFFECT'
      ? Math.expm1(periods * Math.log1p(a / periods))
      : periods * Math.expm1(Math.log1p(a) / periods);
  }
  const due = Number(e !== 0);
  if (name === 'PMT') return payment(a, b, c, d, due);
  if (name === 'FV') return futureValue(a, b, c, d, due);
  if (name === 'PV') {
    const { power, annuity } = growth(a, b);
    return power ? -(d + c * (1 + a * due) * annuity) / power : fail('#DIV/0!');
  }
  if (name === 'NPER') {
    if (a === 0) return b === 0 ? fail('#DIV/0!') : -(c + d) / b;
    if (a <= -1) return fail();
    const pmt = b * (1 + a * due),
      divisor = pmt + a * c;
    const ratio = (pmt - a * d) / divisor;
    return divisor && ratio > 0 ? Math.log(ratio) / Math.log1p(a) : fail();
  }
  // IPMT/PPMT: rate, current period, total periods, present value, future value, timing.
  if (b < 1 || b > c) return fail();
  const timing = Number(f !== 0),
    pmt = payment(a, c, d, e, timing);
  let interest = 0;
  if (a !== 0 && !(timing && b === 1)) {
    interest = futureValue(a, b - 1, pmt, d, timing) * a;
    if (timing) interest /= 1 + a;
  }
  return name === 'IPMT' ? interest : pmt - interest;
}
