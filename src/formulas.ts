import type { Cell, Sheet } from './model';
import SSF from 'ssf';
import { formatNumber } from './number-format';
import { criterionMatches, wildcardMatch } from './formula-criteria';
import { statisticalFunctions, statistic } from './statistics';
import { aggregate, aggregateFunctions, type AggregateItem } from './aggregate';
import { financial, financialFunctions } from './financial';
import { FormulaDependencies } from './formula-dependencies';
import { arrayContains, singleArray } from './sheet-arrays';

export type Value = string | number | boolean;
export interface Calculator {
  (sheet: Sheet, ref: string): Value;
  result(sheet: Sheet, ref: string): { value: Value; kind: 'value' | 'error' | 'unsupported' };
  expression(
    sheet: Sheet,
    formula: string,
  ): { value: Value; kind: 'value' | 'error' | 'unsupported' };
  update(sheets: Sheet[]): Calculator;
  diagnostics(): { evaluations: number; cached: number; syntax: number; edges: number };
}
type ReferenceResult = {
  kind: 'reference';
  rows: number;
  cols: number;
  at(row: number, col: number): Value;
  scalar(): Value;
  values(): Value[];
  blank(): boolean;
};
type ArrayResult = {
  kind: 'array';
  rows: number;
  cols: number;
  at(row: number, col: number): Value;
};
type Result = Value | Value[] | ReferenceResult | ArrayResult;
const isReference = (value: Result): value is ReferenceResult =>
  typeof value === 'object' && !Array.isArray(value) && value.kind === 'reference';
const isArray = (value: Result): value is ArrayResult =>
  typeof value === 'object' && !Array.isArray(value) && value.kind === 'array';
const flatten = (value: Result): Value[] => {
  if (isReference(value)) return value.values();
  if (isArray(value)) {
    if (value.rows * value.cols > 100000) throw new Error('#LIMIT!');
    return Array.from({ length: value.rows * value.cols }, (_, i) =>
      value.at(Math.floor(i / value.cols), i % value.cols),
    );
  }
  return [value].flat();
};
type Node =
  | { kind: 'array'; rows: Node[][] }
  | { kind: 'literal'; value: Value }
  | { kind: 'error'; value: string }
  | { kind: 'missing' }
  | { kind: 'name'; name: string; sheet?: string; workbookScope?: boolean }
  | { kind: 'table'; name: string; spec: string }
  | { kind: 'ref'; ref: string; sheet?: string }
  | {
      kind: 'range';
      from: string;
      to: string;
      sheet?: string;
      axes?: 'columns' | 'rows';
      relativeName?: boolean;
    }
  | { kind: 'unary'; op: string; node: Node }
  | { kind: 'binary'; op: string; left: Node; right: Node }
  | { kind: 'call'; name: string; args: Node[] };
export function colName(index: number): string {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26))
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}
export function coordinates(address: string): [number, number] {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(address);
  if (!match) throw new Error('#REF!');
  return [
    Number(match[2]) - 1,
    [...match[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
  ];
}
export function address(row: number, col: number) {
  return `${colName(col)}${row + 1}`;
}
export function rangeAddresses(a: string, b: string): string[] {
  const [ar, ac] = coordinates(a),
    [br, bc] = coordinates(b);
  if ((Math.abs(br - ar) + 1) * (Math.abs(bc - ac) + 1) > 100000) throw new Error('#LIMIT!');
  const result: string[] = [];
  for (let r = Math.min(ar, br); r <= Math.max(ar, br); r++)
    for (let c = Math.min(ac, bc); c <= Math.max(ac, bc); c++) result.push(address(r, c));
  return result;
}
function parse(input: string): Node {
  const tokens: string[] = [];
  const pattern =
    /\s*("(?:[^"]|"")*"|'(?:[^']|'')*'|#(?:N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)|\$[1-9]\d*|\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+|\$?[\p{L}_\\][\p{L}\p{N}_.$\\]*|<>|<=|>=|[+\-*/^&=<>(),:%!{};])/guy;
  let pos = 0;
  while (pos < input.length) {
    // Keep structured selectors intact, including spaces and escaped header characters.
    if (input.slice(pos).trimStart().startsWith('[')) {
      const start = pos + input.slice(pos).search(/\S/);
      let end = start,
        depth = 0;
      for (; end < input.length; end++) {
        if (input[end] === "'" && /[\[\]#'@]/.test(input[end + 1] || '')) {
          end++;
          continue;
        }
        if (input[end] === '[') depth++;
        if (input[end] === ']' && --depth === 0) break;
      }
      if (depth !== 0) throw new Error('#ERROR!');
      tokens.push(input.slice(start, end + 1));
      pos = end + 1;
      continue;
    }
    pattern.lastIndex = pos;
    const m = pattern.exec(input);
    if (!m) {
      if (!input.slice(pos).trim()) break;
      throw new Error('#ERROR!');
    }
    tokens.push(m[1]);
    pos = pattern.lastIndex;
  }
  if (tokens.length > 2000) throw new Error('#LIMIT!');
  let index = 0;
  const peek = () => tokens[index];
  const take = () => tokens[index++];
  const requireToken = (value: string) => {
    if (take() !== value) throw new Error('#ERROR!');
  };
  function primary(): Node {
    const token = take();
    if (!token) throw new Error('#ERROR!');
    if (token === '{') {
      const rows: Node[][] = [[]];
      do {
        const value = primary();
        const constant = (n: Node): boolean =>
          n.kind === 'literal' ||
          n.kind === 'error' ||
          (n.kind === 'unary' && n.node.kind === 'literal' && typeof n.node.value === 'number');
        if (!constant(value)) throw new Error('#ERROR!');
        rows.at(-1)!.push(value);
        if (peek() !== ',' && peek() !== ';') break;
        if (take() === ';') rows.push([]);
      } while (true);
      requireToken('}');
      if (rows.some((r) => r.length !== rows[0].length)) throw new Error('#ERROR!');
      return { kind: 'array', rows };
    }
    // Excel evaluates negation before exponentiation (unlike conventional math notation).
    if (token === '+' || token === '-') return { kind: 'unary', op: token, node: expression(6) };
    if (token === '(') {
      const node = expression(0);
      requireToken(')');
      return node;
    }
    if (token.startsWith('#')) return { kind: 'error', value: token };
    if (token.startsWith('[') && peek() !== '!') return { kind: 'table', name: '', spec: token };
    if (peek()?.startsWith('[')) return { kind: 'table', name: token, spec: take() };
    if (/^\d|^\.\d/.test(token) && peek() !== ':') return { kind: 'literal', value: Number(token) };
    if (token.startsWith('"'))
      return { kind: 'literal', value: token.slice(1, -1).replaceAll('""', '"') };
    if (peek() === '(') {
      take();
      const args: Node[] = [];
      if (peek() !== ')') {
        do {
          args.push(peek() === ',' || peek() === ')' ? { kind: 'missing' } : expression(0));
          if (peek() !== ',') break;
          take();
        } while (true);
      }
      requireToken(')');
      return { kind: 'call', name: token.toUpperCase().replace(/^_XLFN\./, ''), args };
    }
    let sheet: string | undefined,
      ref = token;
    if (peek() === '!') {
      take();
      sheet = token.startsWith("'") ? token.slice(1, -1).replaceAll("''", "'") : token;
      ref = take();
    }
    if (ref?.startsWith('#')) return { kind: 'error', value: ref };
    if (!sheet && /^(TRUE|FALSE)$/i.test(ref))
      return { kind: 'literal', value: ref.toUpperCase() === 'TRUE' };
    if (peek() === ':' && (/^\$?[A-Z]+$/i.test(ref) || /^\$?[1-9]\d*$/.test(ref))) {
      take();
      const to = take()?.toUpperCase();
      ref = ref.toUpperCase();
      if (/^\$?[A-Z]+$/.test(ref) && /^\$?[A-Z]+$/.test(to))
        return { kind: 'range', from: ref + '1', to: to + '1048576', sheet, axes: 'columns' };
      if (/^\$?[1-9]\d*$/.test(ref) && /^\$?[1-9]\d*$/.test(to))
        return { kind: 'range', from: 'A' + ref, to: 'XFD' + to, sheet, axes: 'rows' };
      throw new Error('#REF!');
    }
    if (!/^\$?[a-z]+\$?[1-9]\d*$/i.test(ref)) return { kind: 'name', name: ref, sheet };
    const [refRow, refCol] = coordinates(ref);
    if ((refRow >= 1048576 || refCol >= 16384) && !ref.includes('$') && peek() !== ':')
      return { kind: 'name', name: ref, sheet };
    ref = ref.toUpperCase();
    if (peek() === ':') {
      take();
      const to = take()?.toUpperCase();
      coordinates(to);
      return { kind: 'range', from: ref, to, sheet };
    }
    return { kind: 'ref', ref, sheet };
  }
  const precedence: Record<string, number> = {
    '=': 1,
    '<>': 1,
    '<': 1,
    '>': 1,
    '<=': 1,
    '>=': 1,
    '&': 2,
    '+': 3,
    '-': 3,
    '*': 4,
    '/': 4,
    '^': 5,
  };
  function expression(min: number): Node {
    let left = primary();
    while (peek() === '%') {
      take();
      left = { kind: 'binary', op: '/', left, right: { kind: 'literal', value: 100 } };
    }
    while (peek() && (precedence[peek()] ?? -1) >= min) {
      const op = take(),
        p = precedence[op];
      left = { kind: 'binary', op, left, right: expression(p + 1) };
    }
    return left;
  }
  const node = expression(0);
  if (index !== tokens.length) throw new Error('#ERROR!');
  return node;
}
const scalar = (v: Result): Value =>
  isReference(v) ? v.scalar() : isArray(v) ? v.at(0, 0) : Array.isArray(v) ? (v[0] ?? '') : v;
export interface ReferenceEndpoint {
  row: number;
  col: number;
  absoluteRow: boolean;
  absoluteCol: boolean;
}
export interface StaticFormulaReference {
  sheetId: string;
  kind: 'cell' | 'range' | 'rows' | 'columns';
  from: ReferenceEndpoint;
  to: ReferenceEndpoint;
}
/** Bind literal syntax without reading cells or evaluating conditional branches.
 * Ranges remain rectangles, including full rows/columns. Incomplete bindings must
 * never be treated as an exhaustive dependency list by callers.
 */
export function bindFormulaReferences(sheets: Sheet[], owner: Sheet, value: string) {
  const references: StaticFormulaReference[] = [];
  let complete = true;
  const endpoint = (ref: string): ReferenceEndpoint => {
    const [row, col] = coordinates(ref);
    if (row < 0 || row >= 1048576 || col < 0 || col >= 16384) throw new Error('#REF!');
    return { row, col, absoluteRow: /\$\d/.test(ref), absoluteCol: ref.startsWith('$') };
  };
  const walk = (node: Node) => {
    if (node.kind === 'ref' || node.kind === 'range') {
      const target = node.sheet
        ? sheets.find((s) => s.name.toLowerCase() === node.sheet!.toLowerCase())
        : sheets.find((s) => s.id === owner.id);
      if (!target || node.sheet?.includes('[')) {
        complete = false;
        return;
      }
      references.push({
        sheetId: target.id,
        kind: node.kind === 'ref' ? 'cell' : node.axes || 'range',
        from: endpoint(node.kind === 'ref' ? node.ref : node.from),
        to: endpoint(node.kind === 'ref' ? node.ref : node.to),
      });
    } else if (node.kind === 'call') {
      if (['INDIRECT', 'OFFSET', 'SUBTOTAL'].includes(node.name)) complete = false;
      node.args.forEach(walk);
    } else if (node.kind === 'binary') {
      walk(node.left);
      walk(node.right);
    } else if (node.kind === 'unary') walk(node.node);
    else if (['name', 'table', 'error'].includes(node.kind)) complete = false;
  };
  if (value.startsWith('=')) {
    try {
      walk(parse(value.slice(1)));
    } catch {
      complete = false;
    }
  }
  return { references, complete };
}
export function formulaReferences(value: string): { sheet?: string; ref: string }[] | null {
  if (!value.startsWith('=')) return [];
  try {
    const result: { sheet?: string; ref: string }[] = [];
    const walk = (node: Node) => {
      if (node.kind === 'ref')
        result.push({ sheet: node.sheet, ref: node.ref.replaceAll('$', '') });
      else if (
        node.kind === 'name' ||
        node.kind === 'table' ||
        (node.kind === 'call' && ['INDIRECT', 'SUBTOTAL'].includes(node.name))
      )
        throw new Error('Dynamic dependency');
      else if (node.kind === 'range')
        result.push(
          ...rangeAddresses(node.from, node.to).map((ref) => ({ sheet: node.sheet, ref })),
        );
      else if (node.kind === 'call') node.args.forEach(walk);
      else if (node.kind === 'binary') {
        walk(node.left);
        walk(node.right);
      } else if (node.kind === 'unary') walk(node.node);
    };
    walk(parse(value.slice(1)));
    return result;
  } catch {
    return null;
  }
}
const num = (v: Result): number => {
  const n = Number(scalar(v));
  if (!Number.isFinite(n)) throw new Error('#VALUE!');
  return n;
};
const truth = (v: Result) => {
  const value = scalar(v);
  return typeof value === 'string' ? value.toUpperCase() === 'TRUE' : Boolean(value);
};
const numBoolean = (v: Result) => typeof scalar(v) === 'number' && scalar(v) !== 0;
export function calculator(sheets: Sheet[]) {
  const memo = new Map<string, Value>(),
    visiting = new Set<string>();
  const failures = new Map<string, Error>();
  const syntax = new Map<string, Node>();
  const names = new Set<string>();
  const lookups = new Map<string, number>();
  const dependencies = new FormulaDependencies();
  const lookupOwners = new Set<string>();
  let evaluations = 0;
  const keyOf = (sheet: Sheet, ref: string) =>
    `${sheet.id}:${ref.replaceAll('$', '').toUpperCase()}`;
  const measureExtents = (sheets: Sheet[]) =>
    new Map(
      sheets.map((s) => {
        let rows = 0,
          cols = 0;
        for (const ref of Object.keys(s.cells)) {
          const [r, c] = coordinates(ref);
          rows = Math.max(rows, r + 1);
          cols = Math.max(cols, c + 1);
        }
        return [s.id, { rows, cols }] as const;
      }),
    );
  let extents = measureExtents(sheets);
  const contexts: { sheet: Sheet; ref: string; array?: boolean }[] = [];
  const sourceCell = (sheet: Sheet, ref: string) => {
    const context = contexts.at(-1);
    dependencies.read(context && keyOf(context.sheet, context.ref), keyOf(sheet, ref));
    return sheet.cells[ref.replaceAll('$', '').toUpperCase()];
  };
  let reads = 0;
  const ast = (text: string) => {
    if (!syntax.has(text)) {
      if (syntax.size >= 4096) syntax.delete(syntax.keys().next().value!);
      syntax.set(text, parse(text.replace(/^=/, '')));
    }
    return syntax.get(text)!;
  };
  const targetSheet = (name: string | undefined, sheet: Sheet) => {
    if (name?.includes('[')) throw new Error('#UNSUPPORTED!');
    const target = name ? sheets.find((s) => s.name.toLowerCase() === name.toLowerCase()) : sheet;
    if (!target) throw new Error('#REF!');
    return target;
  };
  const named = (node: Extract<Node, { kind: 'name' }>, sheet: Sheet) => {
    const workbookScope = node.workbookScope || node.sheet === '[0]';
    const target = targetSheet(node.sheet === '[0]' ? undefined : node.sheet, sheet);
    const definitions = target.nameDefinitions?.filter(
      (d) => d.name.toLowerCase() === node.name.toLowerCase(),
    );
    const definition =
      (!workbookScope && definitions?.find((d) => d.scope === 'worksheet')) ||
      definitions?.find((d) => d.scope === 'workbook');
    if (workbookScope && !target.nameDefinitions) throw new Error('#UNSUPPORTED!');
    const text = target.nameDefinitions
      ? definition?.formula
      : Object.entries(target.definedNames || {}).find(
          ([name]) => name.toLowerCase() === node.name.toLowerCase(),
        )?.[1];
    if (text === undefined) {
      if (
        sheets.some((s) => s.tables?.some((t) => t.name.toLowerCase() === node.name.toLowerCase()))
      )
        return {
          node: { kind: 'table', name: node.name, spec: '[#Data]' } as Node,
          sheet: target,
          identity: `table:${node.name.toLowerCase()}`,
        };
      throw new Error('#NAME?');
    }
    // Excel serializes relative defined-name axes from A1, modulo its grid size.
    // COM's selection-dependent RefersTo display is not the saved origin. Only
    // imported/validated metadata establishes this convention for old models.
    const endpoint = (ref: string) => {
      const [r, c] = coordinates(ref);
      if (r >= 1048576 || c >= 16384) throw new Error('#REF!');
      if (/^\$[A-Z]+\$[1-9]\d*$/i.test(ref)) return ref;
      const caller = contexts.at(-1);
      if (definition?.referenceOrigin !== 'ooxml-a1' || !caller) throw new Error('#UNSUPPORTED!');
      const [row, col] = coordinates(caller.ref);
      return address(
        /\$\d/.test(ref) ? r : (r + row) % 1048576,
        ref.startsWith('$') ? c : (c + col) % 16384,
      );
    };
    // A workbook-defined expression binds unqualified names in workbook scope,
    // even when evaluated on a sheet that shadows one of those names.
    const bind = (n: Node): Node => {
      if (n.kind === 'name' && !n.sheet && definition?.scope === 'workbook')
        return { ...n, workbookScope: true };
      if (n.kind === 'ref') return { ...n, ref: endpoint(n.ref) };
      if (n.kind === 'range') {
        // Whole-axis relative names need separate native intersection evidence.
        if (n.axes) {
          const absolute =
            n.axes === 'columns'
              ? n.from.startsWith('$') && n.to.startsWith('$')
              : /\$\d/.test(n.from) && /\$\d/.test(n.to);
          if (!absolute) throw new Error('#UNSUPPORTED!');
          return n;
        }
        return {
          ...n,
          from: endpoint(n.from),
          to: endpoint(n.to),
          relativeName: ![n.from, n.to].every((ref) => /^\$[A-Z]+\$[1-9]\d*$/i.test(ref)),
        };
      }
      if (n.kind === 'call') return { ...n, args: n.args.map(bind) };
      if (n.kind === 'binary') return { ...n, left: bind(n.left), right: bind(n.right) };
      if (n.kind === 'unary') return { ...n, node: bind(n.node) };
      return n;
    };
    return {
      node: bind(ast(text)),
      sheet: target,
      identity: `${definition?.scope === 'workbook' ? 'workbook' : target.sourcePath || target.id}:${node.name.toLowerCase()}`,
    };
  };
  function indirect(args: Node[], sheet: Sheet): Node {
    if (!args.length || args.length > 2) throw new Error('#VALUE!');
    let text = String(scalar(evaluate(args[0], sheet)));
    const a1 =
      args[1] === undefined ||
      truth(evaluate(args[1], sheet)) ||
      numBoolean(evaluate(args[1], sheet));
    if (!a1) {
      const [r, c] = coordinates(contexts.at(-1)!.ref);
      const match =
        /^(?:(.+)!)?(R(?:\[-?\d+\]|\d*)C(?:\[-?\d+\]|\d*))(?::(R(?:\[-?\d+\]|\d*)C(?:\[-?\d+\]|\d*)))?$/i.exec(
          text,
        );
      if (!match) throw new Error('#REF!');
      const convert = (part: string) => {
        const [, row, col] = /^R(\[-?\d+\]|\d*)C(\[-?\d+\]|\d*)$/i.exec(part)!;
        const axis = (v: string, origin: number) =>
          !v ? origin : v.startsWith('[') ? origin + Number(v.slice(1, -1)) : Number(v) - 1;
        const rr = axis(row, r),
          cc = axis(col, c);
        if (rr < 0 || rr >= 1048576 || cc < 0 || cc >= 16384) throw new Error('#REF!');
        return address(rr, cc);
      };
      text =
        (match[1] ? match[1] + '!' : '') +
        convert(match[2]) +
        (match[3] ? ':' + convert(match[3]) : '');
    }
    let node: Node;
    try {
      node = ast(text);
    } catch {
      throw new Error('#REF!');
    }
    if (!['ref', 'range', 'name', 'table'].includes(node.kind)) throw new Error('#REF!');
    let scope = sheet;
    for (let depth = 0; node.kind === 'name'; depth++) {
      if (depth > 64) throw new Error('#LIMIT!');
      try {
        const resolved = named(node, scope);
        node = resolved.node;
        scope = resolved.sheet;
      } catch (error) {
        if (error instanceof Error && error.message === '#NAME?') throw new Error('#REF!');
        throw error;
      }
    }
    if (['literal', 'binary', 'unary', 'missing'].includes(node.kind)) throw new Error('#REF!');
    let range: ReturnType<typeof reference>;
    try {
      range = reference(node, scope);
    } catch (error) {
      if (error instanceof Error && error.message === '#NAME?') throw new Error('#REF!');
      throw error;
    }
    return {
      kind: 'range',
      sheet: range.sheet.name,
      from: address(range.r, range.c),
      to: address(range.r + range.rows - 1, range.c + range.cols - 1),
    };
  }
  function reference(
    node: Node,
    sheet: Sheet,
    depth = 0,
  ): { sheet: Sheet; r: number; c: number; rows: number; cols: number } {
    if (depth > 64) throw new Error('#LIMIT!');
    if (node.kind === 'name') {
      const resolved = named(node, sheet);
      return reference(resolved.node, resolved.sheet, depth + 1);
    }
    if (node.kind === 'call' && node.name === 'INDIRECT')
      return reference(indirect(node.args, sheet), sheet, depth + 1);
    if (node.kind === 'call' && node.name === 'INDEX') {
      if (node.args.length < 2 || node.args.length > 3) throw new Error('#VALUE!');
      const range = reference(node.args[0], sheet, depth + 1);
      let row = Math.trunc(num(evaluate(node.args[1], sheet))),
        col = node.args[2] ? Math.trunc(num(evaluate(node.args[2], sheet))) : 1;
      if (range.rows === 1 && !node.args[2]) {
        col = row;
        row = 1;
      }
      if (row < 0 || col < 0) throw new Error('#VALUE!');
      if (row > range.rows || col > range.cols) throw new Error('#REF!');
      return {
        ...range,
        r: range.r + (row ? row - 1 : 0),
        c: range.c + (col ? col - 1 : 0),
        rows: row ? 1 : range.rows,
        cols: col ? 1 : range.cols,
      };
    }
    if (node.kind === 'table') {
      const origin = contexts.at(-1)!;
      const [row, col] = coordinates(origin.ref);
      const candidates = sheets.flatMap((s) =>
        (s.tables || []).map((table) => ({ sheet: s, table })),
      );
      const found = candidates.find(({ sheet: s, table: t }) => {
        if (node.name) return t.name.toLowerCase() === node.name.toLowerCase();
        const [start, end = start] = t.ref.split(':');
        const [r, c] = coordinates(start),
          [rr, cc] = coordinates(end);
        return s.id === origin.sheet.id && row >= r && row <= rr && col >= c && col <= cc;
      });
      if (!found) throw new Error('#NAME?');
      const { table, sheet: target } = found;
      const [start, end = start] = table.ref.split(':');
      let [r, c] = coordinates(start);
      let [rr, cc] = coordinates(end);
      const spec = node.spec;
      const selectors = [...spec.matchAll(/\[((?:'[^]|[^\[\]])*)\]/g)].map((m) => m[1]);
      const currentRow =
        selectors.some((s) => s.startsWith('@') || s.toLowerCase() === '#this row') ||
        spec.startsWith('[@[');
      const items = selectors.filter((s) => s.startsWith('#')).map((s) => s.toLowerCase());
      if (
        items.some((s) => !['#all', '#data', '#headers', '#totals', '#this row'].includes(s)) ||
        items.length > 1
      )
        throw new Error('#UNSUPPORTED!');
      if (currentRow) {
        if (row < r + table.headerRows || row > rr - table.totalRows) throw new Error('#VALUE!');
        r = rr = row;
      } else if (items[0] === '#headers') {
        if (!table.headerRows) throw new Error('#REF!');
        rr = r;
      } else if (items[0] === '#totals') {
        if (!table.totalRows) throw new Error('#REF!');
        r = rr;
      } else if (items[0] !== '#all') {
        r += table.headerRows;
        rr -= table.totalRows;
      }
      const columns = selectors
        .filter((s) => !s.startsWith('#'))
        .map((s) => s.replace(/^@/, '').replace(/'([\[\]#'@])/g, '$1'));
      if (columns.length) {
        if (columns.length > 2 || (columns.length === 2 && !spec.includes(']:[')))
          throw new Error('#UNSUPPORTED!');
        const indices = columns.map((s) =>
          table.columns.findIndex((c) => c.toLowerCase() === s.toLowerCase()),
        );
        if (indices.some((i) => i < 0)) throw new Error('#REF!');
        cc = c + Math.max(...indices);
        c += Math.min(...indices);
      }
      if (rr < r) throw new Error('#REF!');
      return { sheet: target, r, c, rows: rr - r + 1, cols: cc - c + 1 };
    }
    if (node.kind !== 'ref' && node.kind !== 'range') throw new Error('#UNSUPPORTED!');
    const target = targetSheet(node.sheet, sheet);
    const [r, c] = coordinates(node.kind === 'ref' ? node.ref : node.from);
    const [rr, cc] = coordinates(node.kind === 'ref' ? node.ref : node.to);
    if (Math.min(r, rr, c, cc) < 0 || Math.max(r, rr) >= 1048576 || Math.max(c, cc) >= 16384)
      throw new Error('#REF!');
    return {
      sheet: target,
      r: Math.min(r, rr),
      c: Math.min(c, cc),
      rows: Math.abs(rr - r) + 1,
      cols: Math.abs(cc - c) + 1,
    };
  }
  /** Retain reference identity until its consumer chooses scalar or range semantics. */
  function referenceValue(
    range: ReturnType<typeof reference>,
    scalarBlank = true,
  ): ReferenceResult {
    const caller = contexts.at(-1);
    const location = caller && coordinates(caller.ref);
    const selected = () => {
      if (!location && range.rows * range.cols > 1) throw new Error('#UNSUPPORTED!');
      const row = range.rows === 1 ? range.r : location![0];
      const col = range.cols === 1 ? range.c : location![1];
      if (
        row < range.r ||
        row >= range.r + range.rows ||
        col < range.c ||
        col >= range.c + range.cols
      )
        throw new Error('#VALUE!');
      return address(row, col);
    };
    return {
      kind: 'reference',
      rows: range.rows,
      cols: range.cols,
      at: (row, col) => {
        const ref = address(range.r + row, range.c + col),
          value = cell(range.sheet, ref);
        return value === '' && !sourceCell(range.sheet, ref)?.value ? 0 : value;
      },
      scalar: () => {
        const ref = selected(),
          value = cell(range.sheet, ref);
        return scalarBlank && value === '' && !sourceCell(range.sheet, ref)?.value ? 0 : value;
      },
      blank: () => !sourceCell(range.sheet, selected())?.value,
      values: () => {
        if (range.rows * range.cols > 100000) throw new Error('#LIMIT!');
        return Array.from({ length: range.rows * range.cols }, (_, i) =>
          cell(
            range.sheet,
            address(range.r + Math.floor(i / range.cols), range.c + (i % range.cols)),
          ),
        );
      },
    };
  }
  function cell(sheet: Sheet, ref: string): Value {
    if (++reads > 100000) throw new Error('#LIMIT!');
    const key = keyOf(sheet, ref);
    const source = sourceCell(sheet, ref);
    const array = sheet.arrayFormulas?.find((a) => {
      if (++reads > 100000) throw new Error('#LIMIT!');
      return arrayContains(a.ref, ref);
    });
    if (array && !singleArray(array)) throw new Error('#UNSUPPORTED!');
    if (memo.has(key)) return memo.get(key)!;
    if (failures.has(key)) throw failures.get(key)!;
    if (visiting.has(key)) throw new Error('#CYCLE!');
    if (visiting.size > 150) throw new Error('#LIMIT!');
    const raw = source?.value || '';
    if (source?.dataType === 'text') return raw;
    if (source?.dataType === 'error') throw new Error(raw);
    if (!raw.startsWith('='))
      return raw.startsWith("'")
        ? raw.slice(1)
        : raw.trim() && Number.isFinite(Number(raw))
          ? Number(raw)
          : /^(true|false)$/i.test(raw)
            ? raw.toLowerCase() === 'true'
            : raw;
    visiting.add(key);
    dependencies.forget(key);
    lookupOwners.delete(key);
    evaluations++;
    contexts.push({ sheet, ref, array: !!array });
    try {
      if (array && array.ref !== ref && array.ref !== `${ref}:${ref}`)
        throw new Error('#UNSUPPORTED!');
      const result = evaluate(ast(raw), sheet);
      const value = array && isReference(result) ? result.at(0, 0) : scalar(result);
      if (typeof value === 'number' && !Number.isFinite(value)) throw new Error('#NUM!');
      memo.set(key, value);
      return value;
    } catch (error) {
      if (error instanceof Error && !['#LIMIT!', '#CYCLE!'].includes(error.message))
        failures.set(key, error);
      throw error;
    } finally {
      visiting.delete(key);
      contexts.pop();
    }
  }
  /** Array context is explicit: legacy IF/IFERROR and scalar callers keep their own context. */
  function evaluate(
    node: Node,
    sheet: Sheet,
    scalarBlank = true,
    arrayContext = contexts.at(-1)?.array || false,
    functionArrays = contexts.at(-1)?.array || false,
  ): Result {
    if (node.kind === 'literal') return node.value;
    if (node.kind === 'array' && !arrayContext) throw new Error('#UNSUPPORTED!');
    if (node.kind === 'array')
      return {
        kind: 'array',
        rows: node.rows.length,
        cols: node.rows[0].length,
        at: (r, c) => scalar(evaluate(node.rows[r][c], sheet)),
      };
    if (node.kind === 'missing') return 0;
    if (node.kind === 'error') throw new Error(node.value);
    if (node.kind === 'name') {
      const resolved = named(node, sheet);
      const caller = contexts.at(-1);
      const key = `${resolved.identity}:${caller ? keyOf(caller.sheet, caller.ref) : 'expression'}`;
      if (names.has(key)) throw new Error('#CYCLE!');
      if (names.size > 64) throw new Error('#LIMIT!');
      names.add(key);
      try {
        return evaluate(resolved.node, resolved.sheet, scalarBlank, arrayContext, functionArrays);
      } finally {
        names.delete(key);
      }
    }
    if (node.kind === 'ref' || node.kind === 'range' || node.kind === 'table')
      return referenceValue(reference(node, sheet), scalarBlank);
    const shaped = (value: Result): ArrayResult => {
      if (isReference(value) || isArray(value)) return { ...value, kind: 'array' };
      if (Array.isArray(value)) throw new Error('#UNSUPPORTED!');
      return { kind: 'array', rows: 1, cols: 1, at: () => value };
    };
    const arrayNumber = (value: Value) => {
      if (typeof value === 'string' && !value.trim()) throw new Error('#VALUE!');
      return num(value);
    };
    if (node.kind === 'unary') {
      const value = evaluate(node.node, sheet, scalarBlank, arrayContext, functionArrays);
      if (arrayContext) {
        const source = shaped(value);
        return {
          ...source,
          at: (r, c) => (node.op === '-' ? -1 : 1) * arrayNumber(source.at(r, c)),
        };
      }
      return (node.op === '-' ? -1 : 1) * num(value);
    }
    if (node.kind === 'binary') {
      if (arrayContext) {
        const left = shaped(evaluate(node.left, sheet, scalarBlank, true, functionArrays));
        const right = shaped(evaluate(node.right, sheet, scalarBlank, true, functionArrays));
        const rows = Math.max(left.rows, right.rows),
          cols = Math.max(left.cols, right.cols);
        if (rows * cols > 100000) throw new Error('#LIMIT!');
        const at = (source: ArrayResult, r: number, c: number) => {
          r = source.rows === 1 ? 0 : r;
          c = source.cols === 1 ? 0 : c;
          if (r >= source.rows || c >= source.cols) throw new Error('#N/A');
          return source.at(r, c);
        };
        return {
          kind: 'array',
          rows,
          cols,
          at: (r, c) => {
            if (++reads > 100000) throw new Error('#LIMIT!');
            const a = at(left, r, c),
              b = at(right, r, c);
            if (['+', '-', '*', '/', '^'].includes(node.op)) {
              arrayNumber(a);
              arrayNumber(b);
            }
            return scalar(
              evaluate(
                {
                  ...node,
                  left: { kind: 'literal', value: a },
                  right: { kind: 'literal', value: b },
                },
                sheet,
                true,
                false,
                false,
              ),
            );
          },
        };
      }
      const left = evaluate(node.left, sheet),
        right = evaluate(node.right, sheet);
      const a = scalar(left),
        b = scalar(right);
      switch (node.op) {
        case '+':
          return num(left) + num(right);
        case '-':
          return num(left) - num(right);
        case '*':
          return num(left) * num(right);
        case '/':
          if (num(right) === 0) throw new Error('#DIV/0!');
          return num(left) / num(right);
        case '^':
          return num(left) ** num(right);
        case '&':
          return String(a) + String(b);
        case '=':
          return String(a).toLowerCase() === String(b).toLowerCase();
        case '<>':
          return String(a).toLowerCase() !== String(b).toLowerCase();
        case '>':
          return a > b;
        case '<':
          return a < b;
        case '>=':
          return a >= b;
        case '<=':
          return a <= b;
      }
    }
    if (node.kind !== 'call') throw new Error('#ERROR!');
    const { name, args } = node;
    // CSE is semantic entry intent. SUMPRODUCT's arithmetic context alone does
    // not lift legacy IF/IFERROR, while a literal condition/value array does.
    if (['IF', 'IFERROR', 'IFNA'].includes(name) && (functionArrays || args[0]?.kind === 'array')) {
      if (args.length < 2 || args.length > (name === 'IF' ? 3 : 2)) throw new Error('#VALUE!');
      const deferred = (arg: Node): ArrayResult => {
        try {
          return shaped(evaluate(arg, sheet, true, true, true));
        } catch (error) {
          return {
            kind: 'array',
            rows: 1,
            cols: 1,
            at: () => {
              throw error;
            },
          };
        }
      };
      const first = deferred(args[0]);
      if (name === 'IF' && first.rows === 1 && first.cols === 1)
        return evaluate(
          truth(first.at(0, 0)) ? args[1] : args[2] || { kind: 'literal', value: false },
          sheet,
          true,
          true,
          true,
        );
      const sources = [first, ...args.slice(1).map(deferred)];
      if (name === 'IF' && sources.length === 2) sources.push(shaped(false));
      const rows = Math.max(...sources.map((s) => s.rows)),
        cols = Math.max(...sources.map((s) => s.cols));
      if (rows * cols > 100000) throw new Error('#LIMIT!');
      const at = (source: ArrayResult, r: number, c: number) => {
        r = source.rows === 1 ? 0 : r;
        c = source.cols === 1 ? 0 : c;
        if (r >= source.rows || c >= source.cols) throw new Error('#N/A');
        return source.at(r, c);
      };
      return {
        kind: 'array',
        rows,
        cols,
        at: (r, c) => {
          if (++reads > 100000) throw new Error('#LIMIT!');
          if (name === 'IF') return at(sources[truth(at(sources[0], r, c)) ? 1 : 2], r, c);
          try {
            return at(sources[0], r, c);
          } catch (error) {
            if (
              !(error instanceof Error) ||
              !/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)$/.test(error.message) ||
              (name === 'IFNA' && error.message !== '#N/A')
            )
              throw error;
            return at(sources[1], r, c);
          }
        },
      };
    }
    if (
      functionArrays &&
      !aggregateFunctions.has(name) &&
      !['SUMPRODUCT', 'TRUE', 'FALSE', 'NA', 'INDEX', 'INDIRECT'].includes(name)
    )
      throw new Error('#UNSUPPORTED!');
    if (name === 'SUMPRODUCT') {
      if (!args.length || args.length > 255) throw new Error('#VALUE!');
      const arrays = args.map((arg) => {
        const value = evaluate(arg, sheet, true, true, functionArrays);
        if (!isReference(value) && !isArray(value) && typeof value !== 'number')
          throw new Error('#VALUE!');
        const source = shaped(value);
        if (source.rows * source.cols > 100000) throw new Error('#LIMIT!');
        // Native SUMPRODUCT observes argument errors before checking dimensions.
        const values = Array.from({ length: source.rows * source.cols }, (_, i) => {
          if (++reads > 100000) throw new Error('#LIMIT!');
          return source.at(Math.floor(i / source.cols), i % source.cols);
        });
        return { ...source, at: (r: number, c: number) => values[r * source.cols + c] };
      });
      const { rows, cols } = arrays[0];
      if (arrays.some((a) => a.rows !== rows || a.cols !== cols)) throw new Error('#VALUE!');
      if (rows * cols > 100000) throw new Error('#LIMIT!');
      let sum = 0;
      for (let r = 0; r < rows; r++)
        for (let c = 0; c < cols; c++) {
          let product = 1;
          for (const a of arrays) {
            if (++reads > 100000) throw new Error('#LIMIT!');
            const value = a.at(r, c);
            product *= typeof value === 'number' ? value : 0;
          }
          sum += product;
        }
      return sum;
    }
    if (name === 'TRUE' || name === 'FALSE') {
      if (args.length) throw new Error('#VALUE!');
      return name === 'TRUE';
    }
    if (
      ['ISBLANK', 'ISNUMBER', 'ISTEXT'].includes(name) &&
      args.length === 1 &&
      ['ref', 'range', 'name', 'table'].includes(args[0].kind)
    ) {
      const value = evaluate(args[0], sheet);
      if (isReference(value) && value.blank()) return name === 'ISBLANK';
      if (name === 'ISBLANK') return false;
      try {
        return typeof scalar(value) === (name === 'ISNUMBER' ? 'number' : 'string');
      } catch (error) {
        if (
          !(error instanceof Error) ||
          ['#UNSUPPORTED!', '#LIMIT!', '#CYCLE!', '#ERROR!'].includes(error.message)
        )
          throw error;
        return false;
      }
    }
    if (name === 'COUNTBLANK' && args.length === 1 && args[0].kind === 'ref') {
      const target = targetSheet(args[0].sheet, sheet);
      return cell(target, args[0].ref.replaceAll('$', '')) === '' ? 1 : 0;
    }
    const fromReference = (arg: Node, scope: Sheet, depth = 0): boolean => {
      if (depth > 64) throw new Error('#LIMIT!');
      if (arg.kind === 'name') {
        const resolved = named(arg, scope);
        return fromReference(resolved.node, resolved.sheet, depth + 1);
      }
      return (
        ['ref', 'range', 'table'].includes(arg.kind) ||
        (arg.kind === 'call' && ['INDIRECT', 'INDEX'].includes(arg.name))
      );
    };
    if (aggregateFunctions.has(name)) {
      if (args.length > 255) throw new Error('#VALUE!');
      function* items(): Generator<AggregateItem> {
        for (const arg of args) {
          try {
            if (fromReference(arg, sheet)) {
              const range = reference(arg, sheet),
                extent = extents.get(range.sheet.id)!;
              for (let r = range.r; r < Math.min(range.r + range.rows, extent.rows); r++)
                for (let c = range.c; c < Math.min(range.c + range.cols, extent.cols); c++) {
                  if (++reads > 100000) throw new Error('#LIMIT!');
                  const ref = address(r, c),
                    source = sourceCell(range.sheet, ref);
                  if (!source || source.value === '') continue;
                  try {
                    yield { value: cell(range.sheet, ref), reference: true };
                  } catch (error) {
                    if (!(error instanceof Error)) throw error;
                    yield { value: error, reference: true };
                  }
                }
            } else {
              const value = evaluate(arg, sheet);
              for (const item of flatten(value))
                yield {
                  value: item,
                  reference: Array.isArray(value) || isReference(value) || isArray(value),
                };
            }
          } catch (error) {
            if (!(error instanceof Error)) throw error;
            yield { value: error, reference: false };
          }
        }
      }
      return aggregate(name, items());
    }
    if (financialFunctions.has(name)) {
      return financial(
        name,
        args.map((arg) => {
          const result = evaluate(arg, sheet);
          let value = scalar(result);
          if (value === '' && isReference(result) && result.blank()) value = 0;
          if (value === '' || (['EFFECT', 'NOMINAL'].includes(name) && typeof value === 'boolean'))
            throw new Error('#VALUE!');
          return num(value);
        }),
      );
    }
    if (statisticalFunctions.has(name)) {
      return statistic(
        name,
        args.map((arg) => ({
          values: flatten(evaluate(arg, sheet, false)),
          reference: fromReference(arg, sheet),
        })),
      );
    }
    if (name === 'SUBTOTAL') {
      if (args.length < 2 || args.length > 255) throw new Error('#VALUE!');
      const code = Math.trunc(num(evaluate(args[0], sheet)));
      const operation = code > 100 ? code - 100 : code;
      if (operation < 1 || operation > 11 || (code > 11 && code < 101)) throw new Error('#VALUE!');
      const numbers: number[] = [];
      let count = 0;
      const nested = (n: Node, scope: Sheet, depth = 0): boolean => {
        if (depth > 64) throw new Error('#LIMIT!');
        if (n.kind === 'call')
          return (
            ['SUBTOTAL', 'AGGREGATE'].includes(n.name) ||
            n.args.some((a) => nested(a, scope, depth + 1))
          );
        if (n.kind === 'binary')
          return nested(n.left, scope, depth + 1) || nested(n.right, scope, depth + 1);
        if (n.kind === 'unary') return nested(n.node, scope, depth + 1);
        if (n.kind === 'name') {
          try {
            const resolved = named(n, scope);
            return nested(resolved.node, resolved.sheet, depth + 1);
          } catch (error) {
            if (error instanceof Error && error.message === '#NAME?') return false;
            throw error;
          }
        }
        return false;
      };
      for (const arg of args.slice(1)) {
        if (!['ref', 'range', 'name', 'table', 'call'].includes(arg.kind))
          throw new Error('#VALUE!');
        const range = reference(arg, sheet),
          target = range.sheet;
        const hidden = new Set(target.hiddenRows);
        const filterRows = (target.autoFilters || [])
          .filter((f) => f.columns.length)
          .map((f) => {
            const [start, end = start] = f.ref.split(':');
            return [coordinates(start)[0] + 1, coordinates(end)[0]];
          });
        const extent = extents.get(target.id)!;
        for (let r = range.r; r < Math.min(range.r + range.rows, extent.rows); r++) {
          // Native Excel uses the persisted worksheet filter mode and hidden-row
          // state. Editing a filter input does not implicitly reapply that filter.
          if (
            (code > 100 ||
              target.filterMode ||
              filterRows.some(([first, last]) => r >= first && r <= last)) &&
            hidden.has(r)
          )
            continue;
          for (let c = range.c; c < Math.min(range.c + range.cols, extent.cols); c++) {
            if (++reads > 100000) throw new Error('#LIMIT!');
            const ref = address(r, c),
              source = sourceCell(target, ref);
            if (!source?.value) continue;
            if (
              source.value.startsWith('=') &&
              source.dataType !== 'text' &&
              nested(ast(source.value), target)
            )
              continue;
            let value: Value;
            try {
              value = cell(target, ref);
            } catch (error) {
              if (
                !(error instanceof Error) ||
                !/^#(?:N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)$/.test(error.message) ||
                ![2, 3].includes(operation)
              )
                throw error;
              count++;
              continue;
            }
            count++;
            if (typeof value === 'number') numbers.push(value);
          }
        }
      }
      if (operation === 2) return numbers.length;
      if (operation === 3) return count;
      if (operation === 4) return numbers.length ? Math.max(...numbers) : 0;
      if (operation === 5) return numbers.length ? Math.min(...numbers) : 0;
      if (operation === 6) return numbers.length ? numbers.reduce((a, b) => a * b, 1) : 0;
      const sum = numbers.reduce((a, b) => a + b, 0);
      if (operation === 9) return sum;
      if (!numbers.length) throw new Error('#DIV/0!');
      if (operation === 1) return sum / numbers.length;
      const sample = operation === 7 || operation === 10;
      if (sample && numbers.length < 2) throw new Error('#DIV/0!');
      const mean = sum / numbers.length;
      const variance =
        numbers.reduce((total, value) => total + (value - mean) ** 2, 0) /
        (numbers.length - Number(sample));
      return operation === 7 || operation === 8 ? Math.sqrt(variance) : variance;
    }
    if (name === 'ROW' || name === 'COLUMN' || name === 'ROWS' || name === 'COLUMNS') {
      if (args.length > 1 || (!args.length && (name === 'ROWS' || name === 'COLUMNS')))
        throw new Error('#VALUE!');
      const range = reference(args[0] || { kind: 'ref', ref: contexts.at(-1)!.ref }, sheet);
      return name === 'ROW'
        ? range.r + 1
        : name === 'COLUMN'
          ? range.c + 1
          : name === 'ROWS'
            ? range.rows
            : range.cols;
    }
    if (name === 'INDIRECT') return evaluate(indirect(args, sheet), sheet);
    if (['ISNA', 'ISERROR', 'ISERR', 'IFNA'].includes(name)) {
      if (args.length !== (name === 'IFNA' ? 2 : 1)) throw new Error('#VALUE!');
      try {
        const value = scalar(evaluate(args[0], sheet));
        return name === 'IFNA' ? value : false;
      } catch (error) {
        if (
          !(error instanceof Error) ||
          ['#UNSUPPORTED!', '#LIMIT!', '#CYCLE!', '#ERROR!'].includes(error.message)
        )
          throw error;
        if (name === 'IFNA') {
          if (error.message !== '#N/A') throw error;
          return evaluate(args[1], sheet);
        }
        return name === 'ISNA'
          ? error.message === '#N/A'
          : name === 'ISERR'
            ? error.message !== '#N/A'
            : true;
      }
    }
    if (name === 'NA') {
      if (args.length) throw new Error('#VALUE!');
      throw new Error('#N/A');
    }
    if (name === 'IF') {
      if (args.length < 2 || args.length > 3) throw new Error('#VALUE!');
      return evaluate(
        truth(evaluate(args[0], sheet)) ? args[1] : args[2] || { kind: 'literal', value: false },
        sheet,
      );
    }
    if (name === 'IFERROR') {
      if (args.length !== 2) throw new Error('#VALUE!');
      try {
        return scalar(evaluate(args[0], sheet));
      } catch (error) {
        if (
          error instanceof Error &&
          ['#UNSUPPORTED!', '#LIMIT!', '#CYCLE!'].includes(error.message)
        )
          throw error;
        return evaluate(args[1], sheet);
      }
    }
    const rectangle = (arg: Node) => {
      const range = reference(arg, sheet);
      return {
        ...range,
        at: (row: number, col: number) => cell(range.sheet, address(range.r + row, range.c + col)),
        blank: (row: number, col: number) =>
          !sourceCell(range.sheet, address(range.r + row, range.c + col))?.value,
      };
    };
    const same = (a: Value, b: Value) =>
      typeof a === typeof b &&
      (typeof a === 'string' ? a.toLowerCase() === String(b).toLowerCase() : a === b);
    const wildcard = wildcardMatch;
    if (name === 'XLOOKUP' || name === 'XMATCH') {
      const lookup = name === 'XLOOKUP';
      if (args.length < (lookup ? 3 : 2) || args.length > (lookup ? 6 : 4))
        throw new Error('#VALUE!');
      const wanted = scalar(evaluate(args[0], sheet));
      const range = rectangle(args[1]);
      if (range.rows !== 1 && range.cols !== 1) throw new Error('#VALUE!');
      const option = (index: number, fallback: number) =>
        !args[index] || args[index].kind === 'missing'
          ? fallback
          : num(evaluate(args[index], sheet));
      const mode = option(lookup ? 4 : 2, 0),
        search = option(lookup ? 5 : 3, 1);
      if (![0, -1, 1, 2].includes(mode) || ![1, -1, 2, -2].includes(search))
        throw new Error('#VALUE!');
      if (mode === 2 && Math.abs(search) === 2) throw new Error('#UNSUPPORTED!');
      const length = Math.max(range.rows, range.cols);
      const at = (n: number) => range.at(range.rows === 1 ? 0 : n, range.rows === 1 ? n : 0);
      const compare = (a: Value, b: Value): number => {
        const rank = (v: Value) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
        if (typeof a !== typeof b) return rank(a) - rank(b);
        if (typeof a === 'string') {
          a = a.toLowerCase();
          b = String(b).toLowerCase();
        }
        return a === b ? 0 : a < b ? -1 : 1;
      };
      let found = -1;
      if (Math.abs(search) === 2) {
        let lo = 0,
          hi = length - 1;
        while (lo <= hi) {
          const mid = Math.floor((lo + hi) / 2),
            delta = compare(at(mid), wanted);
          if (!delta) {
            found = mid;
            break;
          }
          if (delta * Math.sign(search) < 0) lo = mid + 1;
          else hi = mid - 1;
        }
        if (found < 0 && mode !== 0) found = mode * Math.sign(search) > 0 ? lo : hi;
        if (found < 0 || found >= length) found = -1;
      } else {
        let best: Value | undefined;
        for (let i = 0; i < length; i++) {
          const n = search === 1 ? i : length - 1 - i,
            value = at(n);
          if (
            same(value, wanted) ||
            (mode === 2 && typeof wanted === 'string' && wildcard(wanted, value))
          ) {
            found = n;
            break;
          }
          const delta = compare(value, wanted);
          if (
            ((mode === -1 && delta < 0) || (mode === 1 && delta > 0)) &&
            (best === undefined ||
              (mode === -1 ? compare(value, best) > 0 : compare(value, best) < 0))
          ) {
            best = value;
            found = n;
          }
        }
      }
      if (found < 0) {
        if (lookup && args[3] && args[3].kind !== 'missing') return evaluate(args[3], sheet);
        throw new Error('#N/A');
      }
      if (!lookup) return found + 1;
      const returns = rectangle(args[2]);
      if (range.rows === 1 ? returns.cols !== range.cols : returns.rows !== range.rows)
        throw new Error('#VALUE!');
      // Dynamic multi-cell spills need a matrix calculation/grid model. Keep their
      // original caches instead of returning just the first result as a false success.
      if (range.rows === 1 ? returns.rows !== 1 : returns.cols !== 1)
        throw new Error('#UNSUPPORTED!');
      const value = returns.at(range.rows === 1 ? 0 : found, range.rows === 1 ? found : 0);
      return value === '' ? 0 : value;
    }
    if (name === 'VLOOKUP' || name === 'HLOOKUP') {
      if (args.length < 3 || args.length > 4) throw new Error('#VALUE!');
      const wanted = scalar(evaluate(args[0], sheet)),
        table = rectangle(args[1]);
      const index = Math.trunc(num(evaluate(args[2], sheet))) - 1,
        horizontal = name === 'HLOOKUP';
      if (index < 0) throw new Error('#VALUE!');
      if (index >= (horizontal ? table.rows : table.cols)) throw new Error('#REF!');
      const approximate =
        args[3] === undefined ||
        truth(evaluate(args[3], sheet)) ||
        numBoolean(evaluate(args[3], sheet));
      let found = -1;
      const lookupKey = JSON.stringify([
        table.sheet.id,
        table.r,
        table.c,
        table.rows,
        table.cols,
        horizontal,
        approximate,
        wanted,
      ]);
      const extent = extents.get(table.sheet.id)!;
      // Beyond the last occupied row/column the reference contains only blank cells.
      // Include one representative blank for scalar lookups and retain full geometry.
      const length = Math.min(
        horizontal ? table.cols : table.rows,
        Math.max(0, horizontal ? extent.cols - table.c : extent.rows - table.r) + 1,
      );
      const cached = lookups.get(lookupKey);
      // Shared lookup shortcuts do not repeat every source read. Until their
      // cached search paths carry dependencies, invalidate their consumers on edits.
      const owner = contexts.at(-1);
      if (owner) lookupOwners.add(keyOf(owner.sheet, owner.ref));
      for (let n = 0; cached === undefined && n < length; n++) {
        const candidate = horizontal ? table.at(0, n) : table.at(n, 0);
        if (
          same(candidate, wanted) ||
          (!approximate && typeof wanted === 'string' && wildcard(wanted, candidate))
        ) {
          found = n;
          break;
        }
        if (approximate && typeof candidate === typeof wanted && candidate <= wanted) found = n;
      }
      if (cached !== undefined) found = cached;
      else lookups.set(lookupKey, found);
      if (found < 0) throw new Error('#N/A');
      const result = horizontal ? table.at(index, found) : table.at(found, index);
      return result === '' ? 0 : result;
    }
    if (name === 'INDEX') return referenceValue(reference(node, sheet));
    if (name === 'MATCH') {
      if (args.length < 2 || args.length > 3) throw new Error('#VALUE!');
      const wanted = scalar(evaluate(args[0], sheet)),
        range = rectangle(args[1]),
        mode = args[2] ? num(evaluate(args[2], sheet)) : 1;
      if (range.rows !== 1 && range.cols !== 1) throw new Error('#N/A');
      let found = -1;
      for (let n = 0; n < Math.max(range.rows, range.cols); n++) {
        const candidate = range.at(range.rows === 1 ? 0 : n, range.rows === 1 ? n : 0);
        if (
          same(candidate, wanted) ||
          (mode === 0 && typeof wanted === 'string' && wildcard(wanted, candidate))
        ) {
          found = n;
          break;
        }
        if (
          mode !== 0 &&
          typeof candidate === typeof wanted &&
          (mode > 0 ? candidate <= wanted : candidate >= wanted)
        )
          found = n;
      }
      if (found < 0) throw new Error('#N/A');
      return found + 1;
    }
    const criteriaValue = (read: () => Value): Value | Error => {
      try {
        return read();
      } catch (error) {
        if (error instanceof Error) return error;
        throw error;
      }
    };
    const criteriaArgument = (arg: Node): Value => {
      const raw = scalar(evaluate(arg, sheet));
      if (
        raw === '' &&
        arg.kind === 'ref' &&
        !sourceCell(targetSheet(arg.sheet, sheet), arg.ref)?.value
      )
        return 0;
      return raw;
    };
    if (['SUMIFS', 'COUNTIFS', 'AVERAGEIFS'].includes(name)) {
      const start = name === 'COUNTIFS' ? 0 : 1;
      if (args.length < start + 2 || (args.length - start) % 2 || args.length > start + 254)
        throw new Error('#VALUE!');
      const values = rectangle(args[0]);
      const criteria = [];
      for (let i = start; i < args.length; i += 2) {
        const range = rectangle(args[i]);
        if (range.rows !== values.rows || range.cols !== values.cols) throw new Error('#VALUE!');
        const criterion = criteriaArgument(args[i + 1]);
        criteria.push({ range, matches: criterionMatches(criterion) });
      }
      let count = 0,
        sum = 0,
        numbers = 0;
      for (let r = 0; r < values.rows; r++)
        for (let c = 0; c < values.cols; c++) {
          if (
            !criteria.every(({ range, matches }) =>
              matches(
                criteriaValue(() => range.at(r, c)),
                range.blank(r, c),
              ),
            )
          )
            continue;
          count++;
          if (name === 'COUNTIFS') continue;
          const value = values.at(r, c);
          if (typeof value === 'number') {
            sum += value;
            numbers++;
          }
        }
      if (name === 'COUNTIFS') return count;
      if (name === 'AVERAGEIFS' && !numbers) throw new Error('#DIV/0!');
      return name === 'SUMIFS' ? sum : sum / numbers;
    }
    if (name === 'COUNTIF' || name === 'SUMIF' || name === 'AVERAGEIF') {
      if (args.length < 2 || args.length > (name === 'COUNTIF' ? 2 : 3)) throw new Error('#VALUE!');
      const range = rectangle(args[0]),
        values = args[2] ? rectangle(args[2]) : range,
        criterion = criteriaArgument(args[1]);
      const matches = criterionMatches(criterion);
      let count = 0,
        sum = 0,
        numbers = 0;
      for (let r = 0; r < range.rows; r++)
        for (let c = 0; c < range.cols; c++)
          if (
            matches(
              criteriaValue(() => range.at(r, c)),
              range.blank(r, c),
            )
          ) {
            count++;
            if (name === 'COUNTIF') continue;
            const v = values.at(r, c);
            if (typeof v === 'number') {
              sum += v;
              numbers++;
            }
          }
      if (name === 'COUNTIF') return count;
      if (name === 'AVERAGEIF' && !numbers) throw new Error('#DIV/0!');
      return name === 'SUMIF' ? sum : sum / numbers;
    }
    const values = args.map((arg) => evaluate(arg, sheet)),
      flat = () => values.flatMap(flatten);
    const first = () => num(values[0] ?? '');
    switch (name) {
      case 'ABS':
        return Math.abs(first());
      case 'INT':
        return Math.floor(first());
      case 'SQRT':
        return Math.sqrt(first());
      case 'ROUND': {
        const scale = 10 ** num(values[1] ?? 0);
        return (
          (Math.sign(first()) * Math.round((Math.abs(first()) + Number.EPSILON) * scale)) / scale
        );
      }
      case 'MOD':
        if (num(values[1] ?? 0) === 0) throw new Error('#DIV/0!');
        return ((first() % num(values[1])) + num(values[1])) % num(values[1]);
      case 'POWER':
        return first() ** num(values[1] ?? 0);
      case 'ROUNDUP': {
        const scale = 10 ** num(values[1] ?? 0);
        return (Math.sign(first()) * Math.ceil(Math.abs(first()) * scale)) / scale;
      }
      case 'ROUNDDOWN': {
        const scale = 10 ** num(values[1] ?? 0);
        return Math.trunc(first() * scale) / scale;
      }
      case 'COUNTBLANK':
        return flat().filter((v) => v === '').length;
      case 'ISNUMBER':
        return typeof scalar(values[0]) === 'number';
      case 'ISTEXT':
        return typeof scalar(values[0]) === 'string';
      case 'EXACT':
        return String(scalar(values[0])) === String(scalar(values[1]));
      case 'MID': {
        const start = num(values[1]),
          length = num(values[2]);
        if (start < 1 || length < 0) throw new Error('#VALUE!');
        return String(scalar(values[0])).slice(start - 1, start - 1 + length);
      }
      case 'TEXT':
        return formatNumber(String(scalar(values[1])), scalar(values[0]), sheet.date1904);
      case 'DATE': {
        let year = first();
        if (year >= 0 && year < 1900) year += 1900;
        const serial = Date.UTC(year, num(values[1]) - 1, num(values[2])) / 86400000 + 25569;
        return (serial < 61 ? serial - 1 : serial) - (sheet.date1904 ? 1462 : 0);
      }
      case 'YEAR':
      case 'MONTH':
      case 'DAY': {
        const date = SSF.parse_date_code(first(), { date1904: sheet.date1904 });
        if (!date) throw new Error('#NUM!');
        return name === 'YEAR' ? date.y : name === 'MONTH' ? date.m : date.d;
      }
      case 'AND':
        return flat().every(truth);
      case 'OR':
        return flat().some(truth);
      case 'NOT':
        return !truth(values[0] ?? false);
      case 'CONCAT':
      case 'CONCATENATE':
        return flat().join('');
      case 'LEN':
        return String(scalar(values[0] ?? '')).length;
      case 'UPPER':
        return String(scalar(values[0] ?? '')).toUpperCase();
      case 'LOWER':
        return String(scalar(values[0] ?? '')).toLowerCase();
      case 'TRIM':
        return String(scalar(values[0] ?? ''))
          .trim()
          .replace(/ +/g, ' ');
      case 'LEFT':
        return String(scalar(values[0] ?? '')).slice(0, num(values[1] ?? 1));
      case 'RIGHT': {
        const count = num(values[1] ?? 1);
        return count === 0 ? '' : String(scalar(values[0] ?? '')).slice(-count);
      }
      default:
        throw new Error('#UNSUPPORTED!');
    }
  }
  const result = (
    sheet: Sheet,
    ref: string,
  ): { value: Value; kind: 'value' | 'error' | 'unsupported' } => {
    try {
      reads = 0;
      return { value: cell(sheet, ref), kind: 'value' };
    } catch (error) {
      const value =
        error instanceof Error && error.message.startsWith('#') ? error.message : '#ERROR!';
      return {
        value,
        kind: ['#UNSUPPORTED!', '#LIMIT!', '#CYCLE!', '#ERROR!'].includes(value)
          ? 'unsupported'
          : 'error',
      };
    }
  };
  const expression = (
    sheet: Sheet,
    formula: string,
  ): { value: Value; kind: 'value' | 'error' | 'unsupported' } => {
    try {
      reads = 0;
      return { value: scalar(evaluate(parse(formula.replace(/^=/, '')), sheet)), kind: 'value' };
    } catch (error) {
      return { value: error instanceof Error ? error.message : '#ERROR!', kind: 'unsupported' };
    }
  };
  /** Apply an immutable workbook revision. Structural/name/table/filter/extent
   * changes conservatively discard all value caches; ordinary cell edits traverse
   * the observed graph. Neither the snapshots nor stored formula caches are mutated.
   */
  const update = (next: Sheet[]): Calculator => {
    if (next !== sheets) {
      const metadata = (s: Sheet) => {
        const { cells: _cells, ...rest } = s;
        return JSON.stringify(rest);
      };
      const nextExtents = measureExtents(next);
      const structural =
        next.length !== sheets.length ||
        next.some(
          (s, i) =>
            !sheets[i] ||
            metadata(s) !== metadata(sheets[i]) ||
            JSON.stringify(nextExtents.get(s.id)) !== JSON.stringify(extents.get(s.id)),
        );
      if (structural || dependencies.overflow) {
        memo.clear();
        failures.clear();
        dependencies.clear();
        lookupOwners.clear();
      } else {
        const changed = new Set<string>();
        next.forEach((s, i) => {
          const old = sheets[i];
          if (s.cells === old.cells) return;
          for (const ref of new Set([...Object.keys(old.cells), ...Object.keys(s.cells)]))
            if (JSON.stringify(old.cells[ref]) !== JSON.stringify(s.cells[ref]))
              changed.add(keyOf(s, ref));
        });
        if (changed.size) {
          for (const owner of lookupOwners) changed.add(owner);
          const dirty = dependencies.affected(changed);
          for (const key of dirty) {
            memo.delete(key);
            failures.delete(key);
            dependencies.forget(key);
            lookupOwners.delete(key);
          }
        }
      }
      lookups.clear();
      extents = nextExtents;
      sheets = next;
    }
    // New callable identity makes cross-sheet invalidation observable to Svelte
    // even if the visible sheet object itself did not change.
    return api();
  };
  const api = (): Calculator =>
    Object.assign((sheet: Sheet, ref: string): Value => result(sheet, ref).value, {
      result,
      expression,
      update,
      diagnostics: () => ({
        evaluations,
        cached: memo.size,
        syntax: syntax.size,
        edges: dependencies.size,
      }),
    });
  return api();
}
export function displayValue(value: Value, cell?: Cell): string {
  if (cell?.numFmt && cell.numFmt !== 'General') {
    try {
      return formatNumber(cell.numFmt, value, cell.date1904);
    } catch {
      /* Render raw value when a format is unsupported. */
    }
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value !== 'number') return value;
  if (cell?.format === 'currency')
    return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(value);
  if (cell?.format === 'percent')
    return new Intl.NumberFormat('en-US', { style: 'percent', maximumFractionDigits: 2 }).format(
      value,
    );
  if (cell?.format === 'number') return value.toFixed(2);
  return Number(value.toPrecision(12)).toString();
}
export function translateFormula(value: string, dr: number, dc: number): string {
  if (!value.startsWith('=')) return value;
  return value.replace(
    /"(?:[^"]|"")*"|'(?:[^']|'')*'|\[(?:'[^]|[^\[\]])*\]|(?<![\p{L}\p{N}_.\\])(\$?)([A-Z]+)(\$?)([1-9]\d*)(?![\p{L}\p{N}_.\\!([])/giu,
    (match, colLock: string | undefined, col: string, rowLock: string, row: string) => {
      if (colLock === undefined) return match;
      const [r, c] = coordinates(`${col}${row}`),
        nr = r + (rowLock ? 0 : dr),
        nc = c + (colLock ? 0 : dc);
      if (r >= 1048576 || c >= 16384) return match;
      return nr < 0 || nc < 0 || nr >= 1048576 || nc >= 16384
        ? '#REF!'
        : `${colLock}${colName(nc)}${rowLock}${nr + 1}`;
    },
  );
}
