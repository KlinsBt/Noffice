import type { Sheet } from './model';
import { address, coordinates } from './formulas';

export function sheetLayout(sheet: Sheet) {
  const hiddenRows = new Set(sheet.hiddenRows),
    hiddenColumns = new Set(sheet.hiddenColumns);
  const rows = Array.from({ length: sheet.rows }, (_, i) => i).filter((i) => !hiddenRows.has(i));
  const columns = Array.from({ length: sheet.cols }, (_, i) => i).filter(
    (i) => !hiddenColumns.has(i),
  );
  const height = (r: number) => sheet.rowHeights?.[r] || sheet.defaultRowHeight || 30;
  const width = (c: number) => sheet.columnWidths?.[c] || sheet.defaultColumnWidth || 128;
  const offsets = [0];
  for (const row of rows) offsets.push(offsets[offsets.length - 1] + height(row));
  const columnOffsets = [46];
  for (const col of columns)
    columnOffsets.push(columnOffsets[columnOffsets.length - 1] + width(col));
  const merges = (sheet.merges || []).map((ref) => {
    const [a, b = a] = ref.split(':');
    const [r, c] = coordinates(a),
      [rr, cc] = coordinates(b);
    return { r, c, rr, cc, ref: a };
  });
  function cell(r: number, c: number) {
    const merge = merges.find((m) => r >= m.r && r <= m.rr && c >= m.c && c <= m.cc);
    if (!merge)
      return { master: address(r, c), hidden: false, rowspan: 1, colspan: 1, height: height(r) };
    const includedRows = rows.filter((v) => v >= merge.r && v <= merge.rr),
      includedCols = columns.filter((v) => v >= merge.c && v <= merge.cc);
    return {
      master: merge.ref,
      hidden: r !== includedRows[0] || c !== includedCols[0],
      rowspan: includedRows.length,
      colspan: includedCols.length,
      height: includedRows.reduce((sum, v) => sum + height(v), 0),
    };
  }
  function window(scroll: number, viewportHeight = 900) {
    const frozenCount = rows.filter((r) => r < (sheet.frozenRows || 0)).length;
    // Frozen rows remain mounted; virtualize the body behind them using its visible offset.
    const target = Math.min(offsets.at(-1)!, Math.max(0, scroll) + offsets[frozenCount]);
    let low = 0,
      high = offsets.length;
    while (low < high) {
      const mid = (low + high) >>> 1;
      if (offsets[mid] <= target) low = mid + 1;
      else high = mid;
    }
    let start = Math.max(frozenCount, Math.min(rows.length - 1, low - 5));
    let end = Math.min(rows.length, start + 50);
    const visibleEnd = Math.max(0, scroll) + Math.max(0, viewportHeight);
    while (end < rows.length && offsets[end] < visibleEnd) end++;
    end = Math.min(rows.length, end + 4);
    // Keep merged anchors in the DOM when any of their visible rows intersects the window.
    for (const m of merges)
      if (rows[start] >= m.r && rows[start] <= m.rr)
        start = Math.max(
          0,
          rows.findIndex((r) => r >= m.r),
        );
    for (const m of merges)
      if (rows[end - 1] >= m.r && rows[end - 1] <= m.rr)
        end = Math.min(rows.length, rows.findLastIndex((r) => r <= m.rr) + 1);
    const prefix = Math.min(frozenCount, start);
    return {
      start,
      end,
      rows: [...rows.slice(0, prefix), ...rows.slice(start, end)],
      bodyStart: rows[start],
      top: offsets[start] - offsets[prefix],
      bottom: offsets[rows.length] - offsets[end],
    };
  }
  return { rows, columns, height, width, offsets, columnOffsets, cell, window };
}
