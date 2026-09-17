import type { Sheet, WorkbookContent } from './model';
import { coordinates } from './formulas';
import { validationError } from './cell-validation';

export function arrayContains(ref: string, cell: string) {
  const parts = ref.replaceAll('$', '').toUpperCase().split(':');
  if (parts.length > 2 || parts.some((p) => !/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(p)))
    throw Error('Invalid array formula range.');
  const [r, c] = coordinates(cell),
    [a, b] = coordinates(parts[0]),
    [x, y] = coordinates(parts[1] || parts[0]);
  if (x < a || y < b || x >= 1048576 || y >= 16384) throw Error('Invalid array formula range.');
  return r >= a && r <= x && c >= b && c <= y;
}
export function singleArray(array: { anchor: string; ref: string }) {
  return array.ref === array.anchor || array.ref === `${array.anchor}:${array.anchor}`;
}
/** Validate every command path, including paste, clear, sort and retained export. */
export function arrayEditError(before: Sheet, after: Sheet): string | null {
  for (const array of after.arrayFormulas || []) {
    if (
      !arrayContains(array.ref, array.anchor) ||
      !after.cells[array.anchor]?.value.startsWith('=')
    )
      return 'An array formula needs a valid anchor formula.';
    if (
      !singleArray(array) &&
      !before.arrayFormulas?.some((a) => a.anchor === array.anchor && a.ref === array.ref)
    )
      return 'Creating multi-cell array formulas is not supported yet.';
  }
  for (const array of before.arrayFormulas || []) {
    if (singleArray(array)) continue;
    if (!after.arrayFormulas?.some((a) => a.anchor === array.anchor && a.ref === array.ref))
      return 'Changing part or all of a multi-cell array requires an array-aware operation.';
    for (const ref of new Set([...Object.keys(before.cells), ...Object.keys(after.cells)])) {
      if (
        arrayContains(array.ref, ref) &&
        (before.cells[ref]?.value !== after.cells[ref]?.value ||
          before.cells[ref]?.dataType !== after.cells[ref]?.dataType)
      )
        return 'Changing part or all of a multi-cell array requires an array-aware operation.';
    }
  }
  return null;
}
export function enterArrayFormula(
  book: WorkbookContent,
  sheetId: string,
  ref: string,
  value: string,
  array: boolean,
): WorkbookContent {
  const sheet = book.sheets.find((s) => s.id === sheetId)!;
  if (sheet.protected && sheet.cells[ref]?.locked !== false)
    throw Error('This worksheet is protected. Edit an unlocked input cell.');
  if (array && (!value.startsWith('=') || !value.slice(1).trim()))
    throw Error('Enter a formula before using Ctrl+Shift+Enter.');
  const invalid = validationError(
    sheet.cells[ref] || { value: '' },
    value,
    sheet,
    book.sheets,
    ref,
  );
  if (invalid) throw Error(`${ref}: ${invalid}`);
  if (array && sheet.tables?.length) {
    if (sheet.tables.some((t) => arrayContains(t.ref, ref)))
      throw Error('Array entry inside a table is not supported.');
  }
  const next = {
    ...sheet,
    cells: {
      ...sheet.cells,
      [ref]: { ...sheet.cells[ref], value, dataType: undefined, cachedValue: undefined },
    },
    arrayFormulas: [
      ...(sheet.arrayFormulas || []).filter((a) => a.anchor !== ref),
      ...(array ? [{ anchor: ref, ref }] : []),
    ],
  };
  const error = arrayEditError(sheet, next);
  if (error) throw Error(error);
  return { ...book, sheets: book.sheets.map((s) => (s.id === sheetId ? next : s)) };
}
