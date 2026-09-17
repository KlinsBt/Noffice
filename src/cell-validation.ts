import type { Cell, Sheet } from './model';
import { calculator, rangeAddresses } from './formulas';
import { cellValidation } from './sheet-validation';
import type { ValidationRule } from './sheet-validation';

export function validationSetupError(rule: ValidationRule, sheet: Sheet, sheets: Sheet[]) {
  if (rule.type === 'none') return null;
  if (rule.type === 'list')
    return validationChoices({ value: '', validation: rule }, sheet, sheets)
      ? null
      : 'Use a quoted list or a resolvable cell range or defined name.';
  const calc = calculator(sheets);
  const bounds = rule.formulae.map((f) => calc.expression(sheet, String(f)));
  if (bounds.some((b) => b.kind !== 'value' || !Number.isFinite(Number(b.value))))
    return 'Each validation bound must be a number or a supported formula returning a number.';
  if (
    ['between', 'notBetween'].includes(rule.operator || 'between') &&
    Number(bounds[0]?.value) > Number(bounds[1]?.value)
  )
    return 'The minimum must not exceed the maximum.';
  return null;
}

export function validationChoices(cell: Cell, sheet: Sheet, sheets: Sheet[]): string[] | null {
  if (cell.validation?.type !== 'list') return null;
  let source = String(cell.validation.formulae[0] ?? '').replace(/^=/, '');
  for (let depth = 0; depth < 10 && sheet.definedNames?.[source.toLowerCase()]; depth++)
    source = sheet.definedNames[source.toLowerCase()].replace(/^=/, '');
  if (source.startsWith('"') && source.endsWith('"')) return source.slice(1, -1).split(',');
  const match =
    /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?$/i.exec(source);
  if (!match) return null;
  const name = match[1]?.replaceAll("''", "'") || match[2];
  const target = name ? sheets.find((s) => s.name.toLowerCase() === name.toLowerCase()) : sheet;
  if (!target) return null;
  const calc = calculator(sheets);
  try {
    return [
      ...new Set(
        rangeAddresses(match[3], match[4] || match[3])
          .map((ref) => String(calc(target, ref)))
          .filter((v) => v !== ''),
      ),
    ];
  } catch {
    return null;
  }
}
export function validationError(
  cell: Cell,
  value: string,
  sheet: Sheet,
  sheets: Sheet[],
  ref?: string,
): string | null {
  const rule = cell.validation || (ref ? cellValidation(sheet, ref) : undefined);
  if (
    !rule ||
    rule.type === 'none' ||
    !rule.showErrorMessage ||
    (rule.errorStyle && rule.errorStyle.toLowerCase() !== 'stop') ||
    (value === '' && rule.allowBlank)
  )
    return null;
  let valid = true;
  if (rule.type === 'list') {
    const choices = validationChoices({ ...cell, validation: rule }, sheet, sheets);
    if (!choices) return null;
    valid = choices.some((v) => v.toLowerCase() === value.toLowerCase());
  } else if (['whole', 'decimal', 'date', 'time', 'textLength'].includes(rule.type)) {
    const calc = calculator(sheets);
    const scalar = value.startsWith('=')
      ? ref && sheet.cells[ref]?.value === value
        ? calc.result(sheet, ref)
        : calc.expression(sheet, value)
      : { value: Number(value), kind: 'value' };
    const n = rule.type === 'textLength' ? value.length : Number(scalar.value);
    const bounds = rule.formulae.map((f) => calc.expression(sheet, String(f)));
    if (scalar.kind !== 'value' || bounds.some((b) => b.kind !== 'value')) return null;
    const [a, b] = bounds.map((b) => Number(b.value));
    if (!Number.isFinite(n) || (rule.type === 'whole' && !Number.isInteger(n))) valid = false;
    else if (Number.isFinite(a))
      switch (rule.operator || 'between') {
        case 'greaterThan':
          valid = n > a;
          break;
        case 'greaterThanOrEqual':
          valid = n >= a;
          break;
        case 'lessThan':
          valid = n < a;
          break;
        case 'lessThanOrEqual':
          valid = n <= a;
          break;
        case 'equal':
          valid = n === a;
          break;
        case 'notEqual':
          valid = n !== a;
          break;
        case 'notBetween':
          valid = n < a || n > b;
          break;
        default:
          valid = n >= a && n <= b;
      }
  }
  return valid ? null : rule.error || 'This value does not meet the cell’s validation rule.';
}
