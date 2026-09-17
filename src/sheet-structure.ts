import { colName, coordinates } from './formulas';

/** Indices are zero based. Insertion is before `at`; deletion includes `at`. */
export interface StructureEdit {
  sheet: string;
  axis: 'row' | 'column';
  action: 'insert' | 'delete';
  at: number;
  count: number;
}

export function checkedStructureEdit(edit: StructureEdit): StructureEdit {
  const limit = edit.axis === 'row' ? 1048576 : 16384;
  if (
    !edit.sheet ||
    !['row', 'column'].includes(edit.axis) ||
    !['insert', 'delete'].includes(edit.action) ||
    !Number.isInteger(edit.at) ||
    !Number.isInteger(edit.count) ||
    edit.at < 0 ||
    edit.count < 1 ||
    edit.count > 10000 ||
    edit.at + edit.count > limit
  )
    throw Error('Invalid row or column operation.');
  return edit;
}

export function moveIndex(index: number, edit: StructureEdit): number | null {
  if (index < edit.at) return index;
  if (edit.action === 'insert') return index + edit.count;
  return index < edit.at + edit.count ? null : index - edit.count;
}

function moveInterval(
  a: number,
  b: number,
  edit: StructureEdit,
  max: number,
): [number, number] | null {
  if (a > b) {
    const result = moveInterval(b, a, edit, max);
    return result && [result[1], result[0]];
  }
  if (edit.action === 'insert') {
    if (a >= edit.at) a += edit.count;
    if (b >= edit.at) b += edit.count;
    return a >= max ? null : [a, Math.min(b, max - 1)];
  }
  const end = edit.at + edit.count;
  if (b < edit.at) return [a, b];
  if (a >= end) return [a - edit.count, b - edit.count];
  if (a >= edit.at && b < end) return null;
  return [a < edit.at ? a : edit.at, b >= end ? b - edit.count : edit.at - 1];
}

type Endpoint = { kind: 'cell' | 'row' | 'column'; r: number; c: number; rl: string; cl: string };
function endpoint(text: string): Endpoint | null {
  const cell = /^(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6})$/i.exec(text);
  if (cell) {
    const [r, c] = coordinates(text);
    return r < 1048576 && c < 16384 ? { kind: 'cell', r, c, rl: cell[3], cl: cell[1] } : null;
  }
  const column = /^(\$?)([A-Z]{1,3})$/i.exec(text);
  if (column) {
    const c = coordinates(column[2] + '1')[1];
    return c < 16384 ? { kind: 'column', r: 0, c, rl: '', cl: column[1] } : null;
  }
  const row = /^(\$?)([1-9]\d{0,6})$/.exec(text);
  return row && Number(row[2]) <= 1048576
    ? { kind: 'row', r: Number(row[2]) - 1, c: 0, rl: row[1], cl: '' }
    : null;
}
function render(p: Endpoint) {
  return (
    (p.kind !== 'row' ? p.cl + colName(p.c) : '') + (p.kind !== 'column' ? p.rl + (p.r + 1) : '')
  );
}

/** Repairs structural references, including absolute ones; this is not copy/fill translation. */
export function structuralRange(ref: string, edit: StructureEdit): string {
  const parts = ref.split(':').map((p) => p.trim()),
    a = endpoint(parts[0]),
    b = endpoint(parts[1] || parts[0]);
  if (
    !a ||
    !b ||
    parts.length > 2 ||
    a.kind !== b.kind ||
    (parts.length === 1 && a.kind !== 'cell')
  )
    throw Error(`Unsupported worksheet reference: ${ref}`);
  if ((a.kind === 'row' && edit.axis === 'column') || (a.kind === 'column' && edit.axis === 'row'))
    return ref;
  const key = edit.axis === 'row' ? 'r' : 'c',
    max = edit.axis === 'row' ? 1048576 : 16384;
  const moved = moveInterval(a[key], b[key], edit, max);
  if (!moved) return '#REF!';
  a[key] = moved[0];
  b[key] = moved[1];
  return render(a) + (parts.length === 2 ? ':' + render(b) : '');
}

const point = String.raw`\$?[A-Z]{1,3}\$?[1-9]\d{0,6}`;
const reference = `${point}(?:\\s*:\\s*${point})?|\\$?[A-Z]{1,3}\\s*:\\s*\\$?[A-Z]{1,3}|\\$?[1-9]\\d{0,6}\\s*:\\s*\\$?[1-9]\\d{0,6}`;
const sheetName = String.raw`(?:'(?:[^']|'')+'|[\p{L}_\\][\p{L}\p{N}_.\\]*)\s*!\s*`;
const tokens = new RegExp(
  String.raw`"(?:[^"]|"")*"|(?<![\p{L}\p{N}_.\]!])(${sheetName})?(${reference})(?![\p{L}\p{N}_.(])(?!\s*\()`,
  'giu',
);

export function structuralFormula(
  formula: string,
  contextSheet: string,
  edit: StructureEdit,
): string {
  // These need workbook/table-aware parsing; fail atomically rather than partly repair them.
  const unquoted = formula.replace(/"(?:[^"]|"")*"/g, '');
  if (/\[|(?:'[^']*:[^']*'|[\p{L}\p{N}_]+:[\p{L}\p{N}_]+)!/u.test(unquoted))
    throw Error(
      'Row and column operations do not yet support external, table or 3-D formula references.',
    );
  if (new RegExp(`:\\s*${sheetName}`, 'u').test(unquoted))
    throw Error(
      'Ranges with separately qualified endpoints cannot yet be repaired by row and column operations.',
    );
  return formula.replace(tokens, (match, prefix: string | undefined, ref: string | undefined) => {
    if (!ref) return match;
    const name = prefix
      ? prefix.trimEnd().slice(0, -1).trimEnd().replace(/^'|'$/g, '').replace(/''/g, "'")
      : contextSheet;
    if (name.toLocaleLowerCase() !== edit.sheet.toLocaleLowerCase()) return match;
    // Names beyond the A1 grid are not cell references.
    if (!endpoint(ref.split(':')[0].trim()) || !endpoint(ref.split(':').at(-1)!.trim()))
      return match;
    return (prefix || '') + structuralRange(ref, edit);
  });
}
