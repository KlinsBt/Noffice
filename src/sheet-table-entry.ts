import type { WorkbookContent } from './model';
import { address, coordinates } from './formulas';
import { defaultTableStyle, editSheetTable, tableBounds } from './sheet-tables';
import { validationError } from './cell-validation';

/** Direct entry in the first row below an eligible table; never implicit insertion or paste. */
export function tableRowEntry(
  content: WorkbookContent,
  sheetId: string,
  ref: string,
  value: string,
): WorkbookContent | undefined {
  if (!value) return;
  const sheet = content.sheets.find((s) => s.id === sheetId);
  if (!sheet || sheet.protected || sheet.tableEditingBlocked) return;
  const [row, col] = coordinates(ref);
  const candidates = sheet.tables?.filter((t) => {
    const b = tableBounds(t.ref);
    return row === b.rr + 1 && col >= b.c && col <= b.cc;
  });
  if (candidates?.length !== 1) return;
  const table = candidates[0],
    b = tableBounds(table.ref);
  if (value.startsWith('=') && table.totalsActivationBlocked) return;
  for (let c = b.c; c <= b.cc; c++) {
    if (c === col) continue;
    const cell = sheet.cells[address(row, c)];
    if (cell?.value || cell?.dataType === 'text') return;
  }
  const range = `${address(b.r, b.c)}:${address(row, b.cc)}`,
    style = table.style || defaultTableStyle;
  // Unsupported table structures leave this as an ordinary worksheet entry.
  try {
    editSheetTable(content, sheetId, table.name, range, style, true);
  } catch {
    return;
  }
  const entered = {
    ...sheet,
    cells: {
      ...sheet.cells,
      [ref]: { ...sheet.cells[ref], value, dataType: undefined, cachedValue: undefined },
    },
  };
  const proposed = { ...content, sheets: content.sheets.map((s) => (s === sheet ? entered : s)) };
  // Formula entry activates a totals row, excluded from structured data selectors.
  let result = editSheetTable(proposed, sheetId, table.name, range, style, value.startsWith('='));
  if (value.startsWith('=')) {
    if (value.length > 8193 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(value))
      throw Error('Use a totals formula of at most 8,192 characters without control characters.');
    const expanded = result.sheets.find((s) => s.id === sheetId)!,
      cells = { ...expanded.cells };
    const totalFormulas = table.columns.map((_, i) => table.totalFormulas?.[i] ?? null);
    const totalLabels = table.columns.map((_, i) => table.totalLabels?.[i] ?? null);
    totalFormulas[col - b.c] = value.slice(1);
    totalLabels[col - b.c] = null;
    for (let c = b.c; c <= b.cc; c++) {
      if (c === col) continue;
      const key = address(row, c),
        formula = totalFormulas[c - b.c],
        label = totalLabels[c - b.c];
      if (formula)
        cells[key] = {
          ...cells[key],
          value: '=' + formula,
          dataType: undefined,
          cachedValue: undefined,
        };
      else if (label)
        cells[key] = { ...cells[key], value: label, dataType: 'text', cachedValue: undefined };
    }
    const totals = {
      ...expanded,
      cells,
      tables: expanded.tables!.map((t) =>
        t.name === table.name ? { ...t, totalRows: 1, totalFormulas, totalLabels } : t,
      ),
    };
    result = { ...result, sheets: result.sheets.map((s) => (s.id === sheetId ? totals : s)) };
  }
  const next = result.sheets.find((s) => s.id === sheetId)!;
  for (let c = b.c; c <= b.cc; c++) {
    const key = address(row, c),
      cell = next.cells[key];
    if (!cell || (key !== ref && cell.value === sheet.cells[key]?.value)) continue;
    const error = validationError(cell, cell.value, next, result.sheets, key);
    if (error) throw Error(`${key}: ${error}`);
  }
  return result;
}

/** Keep totals labels/formulas synchronized with ordinary edits, paste and deletion. */
export function syncTableTotals(before: WorkbookContent, after: WorkbookContent): WorkbookContent {
  return {
    ...after,
    sheets: after.sheets.map((sheet) => {
      const old = before.sheets.find((s) => s.id === sheet.id);
      if (!old || !sheet.tables?.some((t) => t.totalRows)) return sheet;
      const cells = { ...sheet.cells };
      return {
        ...sheet,
        cells,
        tables: sheet.tables.map((table) => {
          if (!table.totalRows) return table;
          const b = tableBounds(table.ref),
            totalFormulas = table.columns.map((_, i) => table.totalFormulas?.[i] ?? null),
            totalLabels = table.columns.map((_, i) => table.totalLabels?.[i] ?? null);
          for (let c = b.c; c <= b.cc; c++) {
            const key = address(b.rr, c),
              cell = sheet.cells[key];
            if (
              cell?.value === old.cells[key]?.value &&
              cell?.dataType === old.cells[key]?.dataType
            )
              continue;
            const value = cell?.value || '',
              formula = value.startsWith('=') && cell?.dataType !== 'text';
            // Excel stores totals labels as text, including numeric-looking input.
            if (value && !formula)
              cells[key] = { ...cell, dataType: 'text', cachedValue: undefined };
            totalFormulas[c - b.c] = formula ? value.slice(1) : null;
            totalLabels[c - b.c] = formula
              ? null
              : (value.startsWith("'") ? value.slice(1) : value) || null;
          }
          return { ...table, totalFormulas, totalLabels };
        }),
      };
    }),
  };
}
