import type { WorkbookContent } from './model';
import { tableBounds } from './sheet-tables';
import { address } from './formulas';
import type { TableRename } from './table-references';

/** Excel normalizes the entire header row, reserving explicitly supplied suffixes. */
export function normalizedTableHeaders(values: string[], blankPrefix = 'Column') {
  let blank = 1;
  const labels = values.map((text) => (text || `${blankPrefix}${blank++}`).slice(0, 255));
  const reserved = new Set(labels.map((s) => s.toLowerCase())),
    used = new Set<string>();
  return labels.map((label) => {
    let name = label,
      suffix = 2;
    if (used.has(name.toLowerCase()))
      do {
        const tail = String(suffix++);
        name = label.slice(0, 255 - tail.length) + tail;
      } while (reserved.has(name.toLowerCase()) || used.has(name.toLowerCase()));
    used.add(name.toLowerCase());
    return name;
  });
}

/** Keep all nonheader edits, and plan simultaneous column repairs before any state is committed. */
export function prepareHeaderEntry(before: WorkbookContent, proposed: WorkbookContent) {
  const commands: TableRename[] = [];
  const sheets = proposed.sheets.map((s) => {
    const old = before.sheets.find((p) => p.id === s.id);
    if (!old) return s;
    const cells = { ...s.cells };
    for (const table of old.tables || []) {
      if (!table.headerRows) continue;
      const b = tableBounds(table.ref),
        values = [...table.columns];
      let changed = false;
      for (let i = 0; i < table.columns.length; i++) {
        const ref = address(b.r, b.c + i),
          cell = cells[ref];
        if (cell?.value === old.cells[ref]?.value) {
          if (cell && cell.dataType !== old.cells[ref]?.dataType)
            cells[ref] = { ...cell, dataType: old.cells[ref]?.dataType };
          continue;
        }
        if (s.protected || s.tableEditingBlocked)
          throw Error(s.tableEditingBlocked || 'This worksheet is protected.');
        let text = cell?.value || '';
        if (text.startsWith("'")) text = text.slice(1);
        else if (text.startsWith('='))
          throw Error(
            'Table headings are text. Start with an apostrophe to enter a heading beginning with =.',
          );
        if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(text))
          throw Error('This table heading contains an unsupported control character.');
        values[i] = text;
        changed = true;
        // The existing exporter must see consistent headers until the rename patches the package.
        cells[ref] = {
          ...cell,
          value: old.cells[ref]?.value || table.columns[i],
          dataType: old.cells[ref]?.dataType,
        };
      }
      if (!changed) continue;
      const labels = normalizedTableHeaders(values),
        columns = labels.flatMap((name, index) =>
          name === table.columns[index] ? [] : [{ index, name }],
        );
      if (columns.length)
        commands.push({ sheetId: s.id, table: table.name, name: table.name, columns });
    }
    return { ...s, cells };
  });
  return { content: { ...proposed, sheets }, commands };
}

export async function applyHeaderEntry(
  before: WorkbookContent,
  proposed: WorkbookContent,
  rename: (content: WorkbookContent, edit: TableRename) => Promise<WorkbookContent>,
) {
  const plan = prepareHeaderEntry(before, proposed);
  let result = plan.content;
  for (const command of plan.commands) result = await rename(result, command);
  return result;
}
