import { uid, type Sheet, type WorkbookContent, type Cell } from './model';
import { address, coordinates, calculator, displayValue, translateFormula } from './formulas';
import { qualifyTableReferences } from './table-references';
import { validationError } from './cell-validation';
import { arrayContains, singleArray } from './sheet-arrays';

export const legacyCalculatedResizeBlock =
  'Resizing tables with calculated columns, sorting, external data, pivots or slicers is not supported yet.';

export type SheetTable = NonNullable<Sheet['tables']>[number];
export type TableStyle = NonNullable<SheetTable['style']>;
export const tableStyles = [
  'TableStyleMedium2',
  'TableStyleMedium3',
  'TableStyleMedium4',
  'TableStyleMedium5',
  'TableStyleMedium6',
  'TableStyleMedium7',
  '',
];
export const defaultTableStyle: TableStyle = {
  name: tableStyles[0],
  rowStripes: true,
  columnStripes: false,
  firstColumn: false,
  lastColumn: false,
};
export const defaultTableTheme = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646'];

/** Restore additive metadata omitted by older IndexedDB snapshots/backups, keeping all cell edits. */
export function restoreTableMetadata(
  current: WorkbookContent,
  baseline: WorkbookContent,
): WorkbookContent {
  return {
    ...current,
    sheets: current.sheets.map((sheet) => {
      const old = baseline.sheets.find(
        (s) => s.sourcePath === sheet.sourcePath && s.name === sheet.name,
      );
      if (!old) return sheet;
      const nameDefinitions =
        sheet.nameDefinitions?.map((definition) => {
          if (definition.referenceOrigin) return definition;
          const original = old.nameDefinitions?.find(
            (d) =>
              d.name === definition.name &&
              d.scope === definition.scope &&
              d.formula === definition.formula,
          );
          if (!original)
            throw Error(
              'Saved name definitions differ from the source. Preserve a backup before reopening the original workbook.',
            );
          return { ...definition, referenceOrigin: original.referenceOrigin };
        }) ?? old.nameDefinitions;
      if (
        !sheet.nameDefinitions &&
        old.nameDefinitions &&
        JSON.stringify(Object.entries(sheet.definedNames || {}).sort()) !==
          JSON.stringify(Object.entries(old.definedNames || {}).sort())
      )
        throw Error(
          'Saved name definitions differ from the source. Preserve a backup before reopening the original workbook.',
        );
      return {
        ...sheet,
        arrayFormulas:
          sheet.arrayFormulas ??
          (old.arrayFormulas || []).map((array) => {
            const members = singleArray(array)
              ? [array.anchor]
              : [...new Set([...Object.keys(sheet.cells), ...Object.keys(old.cells)])].filter(
                  (ref) => arrayContains(array.ref, ref),
                );
            if (
              members.some(
                (ref) =>
                  sheet.cells[ref]?.value !== old.cells[ref]?.value ||
                  sheet.cells[ref]?.dataType !== old.cells[ref]?.dataType,
              )
            )
              throw Error(
                'Saved array formulas differ from their source. Preserve a backup before reopening the original workbook.',
              );
            return array;
          }),
        nameDefinitions,
        tableTheme: sheet.tableTheme ?? old.tableTheme,
        tableEditingBlocked: sheet.tableEditingBlocked ?? old.tableEditingBlocked,
        tables: sheet.tables?.map((table) => {
          const original = old.tables?.find(
            (t) =>
              (!table.sourcePath || t.sourcePath === table.sourcePath) &&
              t.name === table.name &&
              JSON.stringify(t.columns) === JSON.stringify(table.columns) &&
              t.headerRows === table.headerRows &&
              t.totalRows === table.totalRows,
          );
          return original
            ? {
                ...original,
                ...table,
                sourcePath: original.sourcePath,
                resizeBlocked: original.resizeBlocked,
                style: table.style ?? original.style,
                calculatedColumns: table.calculatedColumns ?? original.calculatedColumns,
                totalLabels: table.totalLabels ?? original.totalLabels,
                totalFormulas: table.totalFormulas ?? original.totalFormulas,
                calculationBlocked: table.calculationBlocked ?? original.calculationBlocked,
              }
            : table.sourcePath && /^xl\/tables\/noffice[a-zA-Z0-9_-]+\.xml$/.test(table.sourcePath)
              ? {
                  ...table,
                  calculatedColumns: table.calculatedColumns ?? table.columns.map(() => null),
                  totalLabels: table.totalLabels ?? table.columns.map(() => null),
                  totalFormulas: table.totalFormulas ?? table.columns.map(() => null),
                }
              : table;
        }),
      };
    }),
  };
}

export function tableBounds(ref: string) {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})(?::\$?([A-Z]{1,3})\$?([1-9]\d{0,6}))?$/i.exec(
    ref.trim(),
  );
  if (!match) throw Error('Enter a cell range such as A1:D12.');
  const [r, c] = coordinates((match[1] + match[2]).toUpperCase());
  const [rr, cc] = coordinates(((match[3] || match[1]) + (match[4] || match[2])).toUpperCase());
  if (rr < r || cc < c || rr >= 1048576 || cc >= 16384)
    throw Error('Use an ordered table range within 10,000 rows, 256 columns and 100,000 cells.');
  return { r, c, rr, cc, ref: `${address(r, c)}:${address(rr, cc)}` };
}
export function tableAt(sheet: Sheet, ref: string) {
  const [r, c] = coordinates(ref);
  return sheet.tables?.find((t) => {
    const b = tableBounds(t.ref);
    return r >= b.r && r <= b.rr && c >= b.c && c <= b.cc;
  });
}
function overlap(a: ReturnType<typeof tableBounds>, b: ReturnType<typeof tableBounds>) {
  return a.r <= b.rr && a.rr >= b.r && a.c <= b.cc && a.cc >= b.c;
}
function editable(sheet: Sheet) {
  if (sheet.protected) throw Error('This worksheet is protected.');
  if (sheet.tableEditingBlocked) throw Error(sheet.tableEditingBlocked);
}
export function nextTableName(sheets: Sheet[]) {
  const names = new Set(
    sheets.flatMap((s) => [
      ...(s.tables || []).map((t) => t.name.toLowerCase()),
      ...Object.keys(s.definedNames || {}).map((n) => n.toLowerCase()),
    ]),
  );
  let i = 1;
  while (names.has(`table${i}`)) i++;
  return `Table${i}`;
}
export function validTableName(name: string, sheets: Sheet[], except?: SheetTable) {
  if (
    !/^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(name) ||
    name.length > 255 ||
    /^[RC]$/i.test(name) ||
    /^R\d*C\d*$/i.test(name)
  )
    throw Error(
      'Use a table name starting with a letter, underscore or backslash, without spaces or cell references.',
    );
  const a1 = /^([A-Z]{1,3})([1-9]\d*)$/i.exec(name);
  if (a1) {
    const [r, c] = coordinates(name.toUpperCase());
    if (r < 1048576 && c < 16384) throw Error('A table name cannot be a cell reference.');
  }
  if (
    sheets.some(
      (s) =>
        s.tables?.some((t) => t !== except && t.name.toLowerCase() === name.toLowerCase()) ||
        Object.keys(s.definedNames || {}).some((n) => n.toLowerCase() === name.toLowerCase()),
    )
  )
    throw Error('This table name is already used in the workbook.');
}
function checkedStyle(style: TableStyle) {
  if (!tableStyles.includes(style.name)) throw Error('Choose one of the supported table styles.');
  return { ...style };
}
function available(sheet: Sheet, ref: string, except?: SheetTable) {
  const b = tableBounds(ref);
  if (b.rr >= 10000 || b.cc >= 256 || (b.rr - b.r + 1) * (b.cc - b.c + 1) > 100000)
    throw Error('Use a table range within 10,000 rows, 256 columns and 100,000 cells.');
  if (b.rr <= b.r) throw Error('Select a header row and at least one data row.');
  if (sheet.tables?.some((t) => t !== except && overlap(b, tableBounds(t.ref))))
    throw Error('Tables cannot overlap.');
  if (sheet.merges?.some((ref) => overlap(b, tableBounds(ref))))
    throw Error('Unmerge cells in the table range first.');
  if (
    sheet.autoFilters?.some(
      (f) => (!except || f.sourcePath !== except.sourcePath) && overlap(b, tableBounds(f.ref)),
    )
  )
    throw Error('This range overlaps an existing filter.');
  return b;
}
export function createSheetTable(
  content: WorkbookContent,
  sheetId: string,
  ref: string,
  name: string,
  style = defaultTableStyle,
): WorkbookContent {
  const sheet = content.sheets.find((s) => s.id === sheetId);
  if (!sheet) throw Error('Select a worksheet.');
  editable(sheet);
  validTableName(name, content.sheets);
  if ((sheet.tables?.length || 0) >= 1000)
    throw Error('This worksheet has reached the table limit.');
  const b = available(sheet, ref),
    cells = { ...sheet.cells },
    columns: string[] = [],
    used = new Set<string>(),
    calc = calculator(content.sheets);
  for (let c = b.c; c <= b.cc; c++) {
    const key = address(b.r, c),
      cell = cells[key];
    let label = cell ? displayValue(calc(sheet, key), cell) : '';
    if (!label) label = `Column${c - b.c + 1}`;
    label = label.slice(0, 255);
    let candidate = label,
      suffix = 2;
    while (used.has(candidate.toLowerCase())) {
      const tail = String(suffix++);
      candidate = label.slice(0, 255 - tail.length) + tail;
    }
    used.add(candidate.toLowerCase());
    columns.push(candidate);
    cells[key] = { ...cell, value: candidate, dataType: 'text', cachedValue: undefined };
  }
  const table: SheetTable = {
    name,
    ref: b.ref,
    columns,
    calculatedColumns: columns.map(() => null),
    totalLabels: columns.map(() => null),
    totalFormulas: columns.map(() => null),
    headerRows: 1,
    totalRows: 0,
    sourcePath: `xl/tables/noffice${uid()}.xml`,
    style: checkedStyle(style),
  };
  const next: Sheet = {
    ...sheet,
    cells,
    rows: Math.max(sheet.rows, b.rr + 1),
    cols: Math.max(sheet.cols, b.cc + 1),
    tables: [...(sheet.tables || []), table],
    autoFilters: [
      ...(sheet.autoFilters || []),
      { ref: b.ref, sourcePath: table.sourcePath, columns: [] },
    ],
  };
  return { ...content, sheets: content.sheets.map((s) => (s.id === sheetId ? next : s)) };
}
export function editSheetTable(
  content: WorkbookContent,
  sheetId: string,
  name: string,
  ref: string,
  style: TableStyle,
  validateOnly = false,
): WorkbookContent {
  const sheet = content.sheets.find((s) => s.id === sheetId),
    table = sheet?.tables?.find((t) => t.name === name);
  if (!sheet || !table) throw Error('The table no longer exists.');
  editable(sheet);
  const b = tableBounds(ref),
    old = tableBounds(table.ref);
  if (b.ref !== old.ref) {
    if (old.rr >= 10000 || old.cc >= 256)
      throw Error('Resizing is limited to tables within 10,000 rows and 256 columns.');
    if (table.resizeBlocked) throw Error(table.resizeBlocked);
    if (table.calculationBlocked) throw Error(table.calculationBlocked);
    if (table.totalRows || !table.headerRows)
      throw Error('Resizing tables with totals or hidden headers is not supported yet.');
    if (b.r !== old.r || b.c !== old.c || b.cc !== old.cc)
      throw Error('Keep the header row and columns fixed when resizing this table.');
    if (sheet.autoFilters?.some((f) => f.columns.length))
      throw Error('Clear filter criteria before resizing the table.');
    available(sheet, b.ref, table);
  }
  // Preserve an imported custom style unless the user explicitly selects another one.
  const nextStyle = style.name === table.style?.name ? { ...style } : checkedStyle(style);
  const next: Sheet = {
    ...sheet,
    rows: Math.max(sheet.rows, b.rr + 1),
    tables: sheet.tables!.map((t) => (t === table ? { ...t, ref: b.ref, style: nextStyle } : t)),
    autoFilters: sheet.autoFilters?.map((f) =>
      f.sourcePath === table.sourcePath ? { ...f, ref: b.ref } : f,
    ),
  };
  const changed: string[] = [];
  if (b.ref !== old.ref && !validateOnly) {
    next.cells = { ...sheet.cells };
    // Extend from the master, preserving exceptions and data already below the table.
    for (let c = b.c; c <= b.cc; c++) {
      const master = table.calculatedColumns?.[c - b.c];
      if (!master) continue;
      for (let r = old.rr + 1; r <= b.rr; r++) {
        const ref = address(r, c),
          cell = next.cells[ref];
        if (cell?.value || cell?.dataType === 'text') continue;
        next.cells[ref] = {
          ...cell,
          value: translateFormula('=' + master, r - (old.r + table.headerRows), 0),
          dataType: undefined,
          cachedValue: undefined,
        };
        changed.push(ref);
      }
    }
    // Excluded rows remain worksheet cells. Their local table selectors need an explicit owner.
    for (let r = b.rr + 1; r <= old.rr; r++)
      for (let c = b.c; c <= b.cc; c++) {
        const ref = address(r, c),
          cell = next.cells[ref];
        if (!cell?.value.startsWith('=') || cell.dataType === 'text') continue;
        const value = qualifyTableReferences(cell.value, table.name);
        next.cells[ref] = { ...cell, value, cachedValue: undefined };
      }
  }
  const result = { ...content, sheets: content.sheets.map((s) => (s === sheet ? next : s)) };
  for (const ref of changed) {
    const error = validationError(next.cells[ref], next.cells[ref].value, next, result.sheets, ref);
    if (error) throw Error(`${ref}: ${error}`);
  }
  return result;
}

/** Ordinary cell edits cannot rename headers without the package reference-repair command. */
export function tableHeaderError(before: Sheet, after: Sheet): string | undefined {
  for (const t of before.tables || []) {
    if (!t.headerRows) continue;
    const b = tableBounds(t.ref);
    for (let c = b.c; c <= b.cc; c++) {
      const key = address(b.r, c);
      if (
        before.cells[key]?.value !== after.cells[key]?.value ||
        before.cells[key]?.dataType !== after.cells[key]?.dataType
      ) {
        const next = after.tables?.find((n) => n.sourcePath === t.sourcePath);
        if (
          next &&
          next.columns[c - b.c] !== t.columns[c - b.c] &&
          after.cells[key]?.value === next.columns[c - b.c] &&
          after.cells[key]?.dataType === 'text'
        )
          continue;
        return 'Use Table design to rename this column and update its structured references.';
      }
    }
  }
}

/** Excel uses integer HLS (0..240) for theme tints, not per-channel RGB blending. */
export function tableStripeColor(hex: string) {
  const [r, g, b] = hex
    .slice(1)
    .match(/../g)!
    .map((s) => parseInt(s, 16));
  const max = Math.max(r, g, b),
    min = Math.min(r, g, b),
    delta = max - min;
  let l = Math.floor(((max + min) * 240 + 255) / 510),
    h = 0,
    s = 0;
  if (delta) {
    s =
      l <= 120
        ? Math.floor((delta * 240 + (max + min) / 2) / (max + min))
        : Math.floor((delta * 240 + (510 - max - min) / 2) / (510 - max - min));
    const d = (c: number) => Math.floor(((max - c) * 40 + delta / 2) / delta);
    h = max === r ? d(b) - d(g) : max === g ? 80 + d(r) - d(b) : 160 + d(g) - d(r);
    if (h < 0) h += 240;
  }
  l = Math.floor(l * 0.2 + 192);
  const m2 =
      l <= 120 ? Math.floor((l * (240 + s) + 120) / 240) : l + s - Math.floor((l * s + 120) / 240),
    m1 = 2 * l - m2;
  const hue = (hue: number) => {
    hue = (hue + 240) % 240;
    return hue < 40
      ? m1 + Math.floor(((m2 - m1) * hue + 20) / 40)
      : hue < 120
        ? m2
        : hue < 160
          ? m1 + Math.floor(((m2 - m1) * (160 - hue) + 20) / 40)
          : m1;
  };
  return (
    '#' +
    [h + 80, h, h - 80]
      .map((v) =>
        Math.floor(((s ? hue(v) : l) * 255 + 120) / 240)
          .toString(16)
          .padStart(2, '0'),
      )
      .join('')
  );
}

/** Built-in medium 2–7 colors use the workbook theme; direct cell formatting wins in the view. */
export function tableCellStyle(sheet: Sheet, ref: string): Partial<Cell> {
  const t = tableAt(sheet, ref),
    style = t?.style;
  if (t && style?.name === '') return { color: '#000000', fill: '#ffffff' };
  if (!t || !style || !/^TableStyleMedium[2-7]$/.test(style.name)) return {};
  const b = tableBounds(t.ref),
    [r, c] = coordinates(ref),
    color = (sheet.tableTheme || defaultTableTheme)[Number(style.name.slice(-1)) - 2];
  const stripe = tableStripeColor(color);
  if (t.headerRows && r === b.r) return { fill: color, color: '#ffffff', bold: true };
  return {
    color: '#000000',
    fill:
      (style.rowStripes && (r - b.r - t.headerRows) % 2 === 0) ||
      (style.columnStripes && (c - b.c) % 2 === 0)
        ? stripe
        : '#ffffff',
    bold:
      (style.firstColumn && c === b.c) ||
      (style.lastColumn && c === b.cc) ||
      (!!t.totalRows && r === b.rr),
  };
}
