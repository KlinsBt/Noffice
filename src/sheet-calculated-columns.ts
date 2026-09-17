import type { Cell, WorkbookContent } from './model';
import { address, coordinates, translateFormula } from './formulas';
import { tableAt, tableBounds } from './sheet-tables';
import { validationError } from './cell-validation';

/** Removing the last body value clears a calculated column's master in the same edit. */
export function clearRemovedColumnFormulas(
  before: WorkbookContent,
  after: WorkbookContent,
): WorkbookContent {
  return {
    ...after,
    sheets: after.sheets.map((sheet) => {
      const old = before.sheets.find((s) => s.id === sheet.id);
      if (!old || !sheet.tables?.some((t) => t.calculatedColumns?.some(Boolean))) return sheet;
      return {
        ...sheet,
        tables: sheet.tables.map((table) => {
          const b = tableBounds(table.ref);
          if (b.rr >= 10000 || b.cc >= 256) return table;
          const formulas = table.calculatedColumns?.map((master, i) => {
            if (!master) return master;
            let removed = false;
            for (let r = b.r + table.headerRows; r <= b.rr - table.totalRows; r++) {
              const ref = address(r, b.c + i),
                cell = sheet.cells[ref];
              if (cell?.value || cell?.dataType === 'text') return master;
              if (old.cells[ref]?.value || old.cells[ref]?.dataType === 'text') removed = true;
            }
            return removed ? null : master;
          });
          return { ...table, calculatedColumns: formulas };
        }),
      };
    }),
  };
}

/** Direct formula entry only. Paste/fill/delete retain their explicit target cells. */
export function calculatedColumnEntry(
  content: WorkbookContent,
  sheetId: string,
  ref: string,
  value: string,
): WorkbookContent | undefined {
  const sheet = content.sheets.find((s) => s.id === sheetId);
  if (!sheet || !value.startsWith('=') || sheet.cells[ref]?.value === value) return;
  const table = tableAt(sheet, ref);
  if (!table) return;
  const b = tableBounds(table.ref),
    [row, col] = coordinates(ref),
    first = b.r + table.headerRows,
    last = b.rr - table.totalRows;
  if (row < first || row > last) return;
  if (last >= 10000 || b.cc >= 256)
    throw Error('Automatic calculated columns are limited to 10,000 rows and 256 columns.');
  const index = col - b.c,
    master = table.calculatedColumns?.[index];
  let empty = true,
    uniform = !!master;
  for (let r = first; r <= last; r++) {
    const key = address(r, col),
      cell = sheet.cells[key];
    if (key !== ref && cell?.value) empty = false;
    if (
      !master ||
      cell?.dataType === 'text' ||
      cell?.value !== translateFormula('=' + master, r - first, 0)
    )
      uniform = false;
  }
  // An existing value elsewhere or any calculated-column exception stops propagation.
  if (!empty && !uniform) return;
  if (sheet.protected || sheet.tableEditingBlocked || table.calculationBlocked)
    throw Error(
      sheet.tableEditingBlocked || table.calculationBlocked || 'This worksheet is protected.',
    );
  if (value.length > 8193 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(value))
    throw Error('Use a formula of at most 8,192 characters without control characters.');
  if (
    sheet.merges?.some((ref) => {
      const m = tableBounds(ref);
      return col >= m.c && col <= m.cc && first <= m.rr && last >= m.r;
    })
  )
    throw Error('Unmerge this table column before entering a calculated formula.');
  const cells: Record<string, Cell> = { ...sheet.cells };
  for (let r = first; r <= last; r++) {
    const key = address(r, col);
    cells[key] = {
      ...cells[key],
      value: translateFormula(value, r - row, 0),
      dataType: undefined,
      cachedValue: undefined,
    };
  }
  const formulas = table.columns.map((_, i) => table.calculatedColumns?.[i] ?? null);
  formulas[index] = translateFormula(value, first - row, 0).slice(1);
  const next = {
    ...sheet,
    cells,
    tables: sheet.tables!.map((t) => (t === table ? { ...t, calculatedColumns: formulas } : t)),
  };
  const sheets = content.sheets.map((s) => (s === sheet ? next : s));
  for (let r = first; r <= last; r++) {
    const key = address(r, col),
      error = validationError(cells[key], cells[key].value, next, sheets, key);
    if (error) throw Error(`${key}: ${error}`);
  }
  return { ...content, sheets };
}
