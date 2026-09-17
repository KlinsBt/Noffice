export interface TableRename {
  sheetId: string;
  table: string;
  name: string;
  column?: { index: number; name: string };
  columns?: { index: number; name: string }[];
}
export interface ReferenceRename {
  table: string;
  name: string;
  column?: { before: string; after: string };
  columns?: { before: string; after: string }[];
}
const identifier = /[\p{L}\p{N}_.$\\]/u;
const same = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();
const escaped = (text: string) => text.replace(/[\[\]#'@]/g, "'$&");

/** Read a complete selector, ignoring Excel's apostrophe-escaped bracket characters. */
function selectorEnd(text: string, start: number) {
  let depth = 0;
  for (let i = start; i < text.length; i++) {
    if (text[i] === "'" && /[\[\]#'@]/.test(text[i + 1] || '')) {
      i++;
      continue;
    }
    if (text[i] === '[') depth++;
    if (text[i] === ']' && --depth === 0) return i + 1;
  }
  throw Error('A structured reference has unbalanced brackets.');
}
function renameSelector(
  spec: string,
  columns: NonNullable<ReferenceRename['columns']>,
  outer = true,
): string {
  let result = '[',
    leaf = '',
    nested = false;
  for (let i = 1; i < spec.length - 1;) {
    if (spec[i] === "'" && /[\[\]#'@]/.test(spec[i + 1] || '')) {
      leaf += spec.slice(i, i + 2);
      i += 2;
    } else if (spec[i] === '[') {
      const end = selectorEnd(spec, i);
      result += leaf + renameSelector(spec.slice(i, end), columns, false);
      leaf = '';
      nested = true;
      i = end;
    } else leaf += spec[i++];
  }
  if (!nested) {
    const row = leaf.startsWith('@') ? '@' : '';
    const raw = leaf.slice(row.length);
    const column = columns.find((c) => same(raw.replace(/'([\[\]#'@])/g, '$1'), c.before));
    if (!raw.startsWith('#') && column) {
      const label = escaped(column.after);
      leaf =
        outer && (/[^\p{L}\p{N} ]/u.test(column.after) || column.after !== column.after.trim())
          ? row + '[' + label + ']'
          : row + label;
    }
  }
  return result + leaf + ']';
}

/** Token repair preserves string literals, other tables, quoted sheet names and function names. */
export function renameTableReferences(
  formula: string,
  edit: ReferenceRename,
  contextTable?: string,
  qualify = false,
) {
  let out = '';
  const columns = edit.columns || (edit.column ? [edit.column] : []);
  for (let i = 0; i < formula.length;) {
    const ch = formula[i];
    if (ch === '"' || ch === "'") {
      const start = i++;
      while (i < formula.length) {
        if (formula[i++] === ch) {
          if (formula[i] === ch) {
            i++;
            continue;
          }
          break;
        }
      }
      if (ch === "'" && formula.slice(start, i).includes('[') && formula[i] === '!')
        throw Error('External workbook references cannot yet be repaired during table renaming.');
      out += formula.slice(start, i);
      continue;
    }
    if (identifier.test(ch)) {
      const start = i++;
      while (i < formula.length && identifier.test(formula[i])) i++;
      const token = formula.slice(start, i);
      // OOXML stores a whole-table reference as Table[#Data]. A bare identifier
      // is a name token, which native Excel does not bind to the table on load.
      const target = same(token, edit.table) && formula[i] === '[';
      out += target ? edit.name : token;
      if (formula[i] === '[') {
        const stop = selectorEnd(formula, i),
          spec = formula.slice(i, stop);
        out += target && columns.length ? renameSelector(spec, columns) : spec;
        i = stop;
      }
      continue;
    }
    if (ch === '[') {
      const end = selectorEnd(formula, i),
        spec = formula.slice(i, end);
      // External workbook prefixes require a separate dependency model.
      if (identifier.test(formula[end] || '') || formula[end] === '!')
        throw Error('External workbook references cannot yet be repaired during table renaming.');
      const changed = columns.length ? renameSelector(spec, columns) : spec;
      if (changed !== spec && !contextTable)
        throw Error('An unqualified column reference has no unambiguous table context.');
      out +=
        (qualify ? edit.name : '') +
        (contextTable && same(contextTable, edit.table) ? changed : spec);
      i = end;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}

/** Keep formerly local selectors bound to their table when rows leave its range. */
export function qualifyTableReferences(formula: string, table: string) {
  return renameTableReferences(formula, { table, name: table }, table, true);
}
