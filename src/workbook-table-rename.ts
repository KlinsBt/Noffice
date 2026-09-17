import type { OfficeFile, WorkbookContent } from './model';
import { exportOffice, inspectZip } from './formats';
import { importWorkbook } from './xlsx-import';
import { validTableName } from './sheet-tables';
import { renameXlsxTable } from './xlsx-table-rename';
import type { TableRename } from './table-references';

/** A single undoable checkpoint, with all pending cell/style edits included before reference repair. */
export async function renameWorkbookTable(
  file: OfficeFile,
  edit: TableRename,
): Promise<WorkbookContent> {
  if (file.content.kind !== 'excel') throw Error('Select a workbook.');
  const previous = file.content,
    sheet = previous.sheets.find((s) => s.id === edit.sheetId);
  const table = sheet?.tables?.find((t) => t.name === edit.table);
  if (!sheet || !table) throw Error('The table no longer exists.');
  if (sheet.protected || sheet.tableEditingBlocked)
    throw Error(sheet.tableEditingBlocked || 'This worksheet is protected.');
  if (edit.name !== table.name) validTableName(edit.name, previous.sheets, table);
  const edits = edit.columns || (edit.column ? [edit.column] : []),
    finalColumns = [...table.columns];
  if (edits.length > 256 || new Set(edits.map((c) => c.index)).size !== edits.length)
    throw Error('Choose each table column at most once.');
  for (const { index, name } of edits) {
    if (!Number.isInteger(index) || index < 0 || index >= table.columns.length)
      throw Error('Select a table column.');
    if (!name.length || name.length > 255 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(name))
      throw Error('Use a nonempty column name of at most 255 characters.');
    finalColumns[index] = name;
  }
  if (new Set(finalColumns.map((n) => n.toLowerCase())).size !== finalColumns.length)
    throw Error('Another column already uses this name.');
  const columns = finalColumns.flatMap((name, index) =>
    name === table.columns[index] ? [] : [{ before: table.columns[index], after: name }],
  );
  if (edit.name === table.name && !columns.length) return previous;
  const input = await (await exportOffice(file)).arrayBuffer();
  inspectZip(input);
  // Newly generated files also have real part identities after export.
  const baseline = await importWorkbook(input);
  const source = baseline.sheets[previous.sheets.indexOf(sheet)]?.tables?.find(
    (t) => t.name === table.name,
  );
  if (!source?.sourcePath) throw Error('The table source could not be located.');
  const data = await renameXlsxTable(input, {
    table: table.name,
    name: edit.name,
    sourcePath: source.sourcePath,
    columns,
  });
  inspectZip(data);
  if (data.byteLength > 50000000) throw Error('Workbook checkpoints are limited to 50 MB.');
  const next = await importWorkbook(data);
  next.sheets.forEach((s, i) => {
    s.id = previous.sheets[i].id;
    s.rows = Math.max(s.rows, previous.sheets[i].rows);
    s.cols = Math.max(s.cols, previous.sheets[i].cols);
  });
  if (!previous.xlsxStructureBase && !file.original?.name.toLowerCase().endsWith('.xlsx')) {
    for (const s of next.sheets) {
      delete s.sourcePath;
      for (const filter of s.autoFilters || [])
        if (!s.tables?.some((t) => t.sourcePath === filter.sourcePath)) delete filter.sourcePath;
    }
    return next;
  }
  let binary = '';
  const bytes = new Uint8Array(data);
  for (let i = 0; i < bytes.length; i += 8192)
    binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
  next.xlsxStructureBase = btoa(binary);
  return next;
}
