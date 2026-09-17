import type { Value } from './formulas';
export const aggregateFunctions = new Set(['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA']);
export type AggregateItem = { value: Value | Error; reference: boolean; blank?: boolean };
/** Preserve argument provenance: an empty cell differs from a formula returning empty text. */
export function aggregate(name: string, items: Iterable<AggregateItem>): number {
  let count = 0,
    sum = 0,
    min = Infinity,
    max = -Infinity;
  for (const { value, reference, blank } of items) {
    if (value instanceof Error) {
      if (
        !/^#(?:N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)$/.test(value.message) ||
        !['COUNT', 'COUNTA'].includes(name)
      )
        throw value;
      if (name === 'COUNTA') count++;
      continue;
    }
    if (blank) continue;
    if (name === 'COUNTA') {
      count++;
      continue;
    }
    let number: number;
    if (typeof value === 'number') number = value;
    else if (reference) continue;
    else if (typeof value === 'boolean') number = Number(value);
    else if (value.trim() && Number.isFinite(Number(value))) number = Number(value);
    else {
      if (name === 'COUNT') continue;
      throw new Error('#VALUE!');
    }
    count++;
    sum += number;
    min = Math.min(min, number);
    max = Math.max(max, number);
  }
  if (name === 'COUNT' || name === 'COUNTA') return count;
  if (name === 'MIN') return count ? min : 0;
  if (name === 'MAX') return count ? max : 0;
  if (name === 'AVERAGE') {
    if (!count) throw new Error('#DIV/0!');
    return sum / count;
  }
  return sum;
}
