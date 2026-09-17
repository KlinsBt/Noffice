import type { Cell, Sheet } from './model';
import { coordinates, translateFormula } from './formulas';
export type ValidationRule = NonNullable<Cell['validation']>;
export const validationTypes = ['whole', 'decimal', 'list', 'textLength'] as const;
export const validationOperators = [
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'lessThan',
  'greaterThanOrEqual',
  'lessThanOrEqual',
] as const;
export function validationBounds(ref: string) {
  const [a, b = a] = ref.split(':');
  const [r, c] = coordinates(a),
    [rr, cc] = coordinates(b);
  if (Math.min(r, c, rr, cc) < 0 || Math.max(r, rr) > 1048575 || Math.max(c, cc) > 16383)
    throw Error('Invalid validation range.');
  return [Math.min(r, rr), Math.min(c, cc), Math.max(r, rr), Math.max(c, cc)];
}
export function shiftValidation(rule: ValidationRule, from: string, to: string): ValidationRule {
  const [r, c] = coordinates(from),
    [rr, cc] = coordinates(to);
  if (r === rr && c === cc) return rule;
  return {
    ...rule,
    formulae: rule.formulae.map((formula) =>
      translateFormula('=' + String(formula).replace(/^=/, ''), rr - r, cc - c).slice(1),
    ),
  };
}
export function cellValidation(sheet: Sheet, ref: string): ValidationRule | undefined {
  if (sheet.cells[ref]?.validation) return sheet.cells[ref].validation;
  const [r, c] = coordinates(ref);
  for (const entry of sheet.validationRanges || []) {
    const refs = entry.ref.trim().split(/\s+/);
    for (const range of refs) {
      const [top, left, bottom, right] = validationBounds(range);
      if (r >= top && r <= bottom && c >= left && c <= right)
        return shiftValidation(entry.rule, refs[0].split(':')[0], ref);
    }
  }
}
export function checkedValidation(rule: ValidationRule) {
  if (rule.type === 'none') return { type: 'none', formulae: [] };
  if (!(validationTypes as readonly string[]).includes(rule.type))
    throw Error('Choose a supported validation type.');
  if (
    rule.type !== 'list' &&
    !(validationOperators as readonly string[]).includes(rule.operator || 'between')
  )
    throw Error('Choose a supported comparison.');
  const count =
    rule.type === 'list'
      ? 1
      : ['between', 'notBetween'].includes(rule.operator || 'between')
        ? 2
        : 1;
  const formulae = rule.formulae.slice(0, count).map((f) => String(f).trim().replace(/^=/, ''));
  if (formulae.length !== count || formulae.some((f) => !f || f.length > 255))
    throw Error('Enter each required value or formula (up to 255 characters).');
  if (rule.type === 'list' && formulae[0].startsWith('"') && !/^"[^"\r\n]*"$/.test(formulae[0]))
    throw Error('Enter a quoted comma-separated list or a cell-range reference.');
  if (
    (rule.prompt?.length || 0) > 255 ||
    (rule.error?.length || 0) > 225 ||
    (rule.promptTitle?.length || 0) > 32 ||
    (rule.errorTitle?.length || 0) > 32
  )
    throw Error('Validation messages or titles are too long.');
  if (rule.errorStyle && rule.errorStyle !== 'stop')
    throw Error('This editor currently supports Stop alerts for new rules.');
  return { ...rule, formulae };
}
export function applyValidation(sheet: Sheet, refs: string[], rule: ValidationRule): Sheet {
  if (sheet.protected) throw Error('Data validation cannot be changed on a protected worksheet.');
  if (!refs.length || refs.length > 10000) throw Error('Select between 1 and 10,000 cells.');
  const valid = checkedValidation(rule),
    cells = { ...sheet.cells };
  let changed = false;
  for (const ref of refs) {
    const next = shiftValidation(valid, refs[0], ref);
    if (JSON.stringify(cellValidation(sheet, ref)) === JSON.stringify(next)) continue;
    cells[ref] = { ...(cells[ref] || { value: '' }), validation: next };
    changed = true;
  }
  return changed ? { ...sheet, cells } : sheet;
}
