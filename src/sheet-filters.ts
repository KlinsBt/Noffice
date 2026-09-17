import type { Sheet, Cell } from './model';
import { criterionMatches } from './formula-criteria';
import { formatNumber } from './number-format';
import { address, calculator, coordinates } from './formulas';

export type SheetFilter = NonNullable<Sheet['autoFilters']>[number];
type Filter = SheetFilter;
const children = (root: Element, name: string) =>
  Array.from(root.children).filter((e) => e.localName === name);

/** Retain only semantics understood by the calculator; mark every other filter explicit. */
export function readAutoFilters(root: Element, sourcePath?: string): Filter[] {
  return children(root, 'autoFilter').map((filter) => ({
    ref: filter.getAttribute('ref') || '',
    ...(sourcePath ? { sourcePath } : {}),
    // Empty filterColumn entries describe dropdown visibility, not active criteria.
    columns: children(filter, 'filterColumn')
      .filter((column) => column.children.length > 0)
      .map((column) => {
        const values = children(column, 'filters')[0],
          custom = children(column, 'customFilters')[0];
        const rules =
          custom &&
          children(custom, 'customFilter').map((rule) => ({
            operator: rule.getAttribute('operator') || 'equal',
            value: rule.getAttribute('val') || '',
          }));
        const supported = new Set([
          'equal',
          'notEqual',
          'lessThan',
          'lessThanOrEqual',
          'greaterThan',
          'greaterThanOrEqual',
        ]);
        return {
          col: Number(column.getAttribute('colId')),
          ...(values
            ? {
                values: children(values, 'filter').map((e) => e.getAttribute('val') || ''),
                blank: values.getAttribute('blank') === '1',
              }
            : {}),
          ...(custom ? { custom: rules, and: custom.getAttribute('and') === '1' } : {}),
          ...(!Number.isInteger(Number(column.getAttribute('colId'))) ||
          column.children.length !== 1 ||
          (!values && !custom) ||
          (values && Array.from(values.children).some((e) => e.localName !== 'filter')) ||
          (custom &&
            (!rules?.length ||
              rules.length > 2 ||
              rules.some((r) => !supported.has(r.operator)) ||
              Array.from(custom.children).some((e) => e.localName !== 'customFilter')))
            ? { unsupported: true }
            : {}),
        };
      }),
  }));
}

/** Applying a filter is an explicit visibility operation; ordinary edits leave it unchanged. */
export function applySheetFilters(sheet: Sheet, sheets: Sheet[], filters: SheetFilter[]): Sheet {
  if (sheet.protected) throw new Error('This worksheet is protected.');
  const hidden = new Set(sheet.hiddenRows),
    calc = calculator(sheets);
  const bounds = (filter: SheetFilter) => {
    const [start, end = start] = filter.ref.split(':');
    const [r, c] = coordinates(start),
      [rr, cc] = coordinates(end);
    if (
      rr <= r ||
      cc < c ||
      rr >= sheet.rows ||
      cc >= sheet.cols ||
      (rr - r + 1) * (cc - c + 1) > 100000
    )
      throw new Error('Select a header and data rows within the supported range size.');
    return { r, c, rr, cc };
  };
  for (const filter of [...(sheet.autoFilters || []), ...filters]) {
    const { r, rr } = bounds(filter);
    for (let row = r + 1; row <= rr; row++) hidden.delete(row);
  }
  let reads = 0;
  for (const filter of filters) {
    const { r, c, rr, cc } = bounds(filter);
    for (const column of filter.columns)
      if (column.unsupported || column.col > cc - c)
        throw new Error(
          'This filter type cannot be reapplied yet. Clear it to choose a supported filter.',
        );
    for (let row = r + 1; row <= rr; row++)
      for (const column of filter.columns) {
        if (++reads > 100000)
          throw new Error('This filter exceeds the supported calculation limit.');
        const ref = address(row, c + column.col),
          result = calc.result(sheet, ref);
        if (result.kind === 'unsupported')
          throw new Error(`Cannot filter ${ref}: its formula is unsupported.`);
        const matches = filterMatches(column, result.value, sheet.cells[ref], sheet.date1904);
        if (matches === undefined) throw new Error('This filter type cannot be reapplied yet.');
        if (!matches) hidden.add(row);
      }
  }
  return {
    ...sheet,
    autoFilters: filters,
    filterMode: filters.some((f) => f.columns.length),
    hiddenRows: [...hidden].sort((a, b) => a - b),
  };
}

export function writeFilterColumns(filter: Element, columns: SheetFilter['columns']) {
  const buttonAttributes = new Map(
    children(filter, 'filterColumn').map((e) => [
      Number(e.getAttribute('colId')),
      ['hiddenButton', 'showButton'].flatMap((name) =>
        e.hasAttribute(name) ? [[name, e.getAttribute(name)!]] : [],
      ),
    ]),
  );
  children(filter, 'filterColumn').forEach((e) => e.remove());
  const create = (name: string) => filter.ownerDocument.createElementNS(filter.namespaceURI, name);
  for (const column of columns) {
    if (column.unsupported) throw new Error('Cannot write an unsupported filter rule.');
    const node = create('filterColumn');
    node.setAttribute('colId', String(column.col));
    for (const [name, value] of buttonAttributes.get(column.col) || [])
      node.setAttribute(name, value);
    buttonAttributes.delete(column.col);
    if (column.values) {
      const values = create('filters');
      if (column.blank) values.setAttribute('blank', '1');
      for (const value of column.values) {
        const entry = create('filter');
        entry.setAttribute('val', value);
        values.appendChild(entry);
      }
      node.appendChild(values);
    } else if (column.custom) {
      const custom = create('customFilters');
      if (column.and) custom.setAttribute('and', '1');
      for (const rule of column.custom) {
        const entry = create('customFilter');
        entry.setAttribute('operator', rule.operator);
        entry.setAttribute('val', rule.value);
        custom.appendChild(entry);
      }
      node.appendChild(custom);
    } else throw new Error('Missing filter criteria.');
    filter.insertBefore(
      node,
      Array.from(filter.children).find((e) => e.localName !== 'filterColumn') || null,
    );
  }
  for (const [col, attrs] of buttonAttributes) {
    if (!attrs.length) continue;
    const node = create('filterColumn');
    node.setAttribute('colId', String(col));
    for (const [name, value] of attrs) node.setAttribute(name, value);
    filter.insertBefore(
      node,
      Array.from(filter.children).find((e) => e.localName !== 'filterColumn') || null,
    );
  }
}

/** Returns undefined for unsupported semantics, allowing callers to preserve caches. */
export function filterMatches(
  column: Filter['columns'][number],
  value: string | number | boolean,
  cell: Cell | undefined,
  date1904 = false,
): boolean | undefined {
  if (column.unsupported) return undefined;
  if (column.values) {
    if (value === '') return !!column.blank;
    const shown =
      typeof value === 'boolean'
        ? value
          ? 'TRUE'
          : 'FALSE'
        : typeof value === 'number' && cell?.numFmt
          ? formatNumber(cell.numFmt, value, date1904)
          : String(value);
    return column.values.some(
      (candidate) => candidate.toLocaleLowerCase('en-US') === shown.toLocaleLowerCase('en-US'),
    );
  }
  if (column.custom?.length) {
    const operators: Record<string, string> = {
      equal: '=',
      notEqual: '<>',
      lessThan: '<',
      lessThanOrEqual: '<=',
      greaterThan: '>',
      greaterThanOrEqual: '>=',
    };
    const results = column.custom.map((rule) => {
      const op = operators[rule.operator];
      if (!op) throw new Error('#UNSUPPORTED!');
      return criterionMatches(op + rule.value)(value, value === '');
    });
    return column.and ? results.every(Boolean) : results.some(Boolean);
  }
  return undefined;
}
