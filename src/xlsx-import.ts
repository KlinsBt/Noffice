import type ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { uid, type Cell, type Sheet, type WorkbookContent } from './model';
import { coordinates } from './formulas';
import { readAutoFilters } from './sheet-filters';
import { readWorksheetLinks } from './xlsx-links';
import { readValidationRanges } from './xlsx-validation';
import { cellValidation } from './sheet-validation';
import { tableBounds } from './sheet-tables';
import { arrayContains } from './sheet-arrays';

export const spreadsheetNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export function parseXML(text: string) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('Invalid workbook XML.');
  return doc;
}
export function elements(root: Document | Element, name: string) {
  return Array.from(root.getElementsByTagNameNS('*', name));
}
export function resolvePart(from: string, target: string) {
  const segments = (
    target.startsWith('/') ? target.slice(1) : from.slice(0, from.lastIndexOf('/') + 1) + target
  ).split('/');
  const result: string[] = [];
  for (const segment of segments) {
    if (segment === '..') result.pop();
    else if (segment && segment !== '.') result.push(segment);
  }
  return result.join('/');
}
export async function workbookSheets(zip: JSZip) {
  const book = parseXML(await zip.file('xl/workbook.xml')!.async('string'));
  const rels = parseXML(await zip.file('xl/_rels/workbook.xml.rels')!.async('string'));
  const paths = new Map(
    elements(rels, 'Relationship')
      .filter((e) => e.getAttribute('TargetMode') !== 'External')
      .map((e) => [
        e.getAttribute('Id'),
        resolvePart('xl/workbook.xml', e.getAttribute('Target')!),
      ]),
  );
  return {
    book,
    sheets: elements(book, 'sheet').map((e) => ({
      name: e.getAttribute('name')!,
      path: paths.get(e.getAttribute('r:id'))!,
      element: e,
    })),
  };
}
function colorResolver(theme: Document | null) {
  const scheme = theme && elements(theme, 'clrScheme')[0];
  const names = [
    'lt1',
    'dk1',
    'lt2',
    'dk2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ];
  const defaults = [
    'FFFFFF',
    '000000',
    'E7E6E6',
    '44546A',
    '4472C4',
    'ED7D31',
    'A5A5A5',
    'FFC000',
    '5B9BD5',
    '70AD47',
    '0563C1',
    '954F72',
  ];
  const indexed = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF'];
  return (color?: Partial<ExcelJS.Color> & { indexed?: number }): string | undefined => {
    if (!color) return undefined;
    let hex = color.argb?.slice(-6);
    if (color.theme !== undefined) {
      const entry = scheme && elements(scheme, names[color.theme])[0]?.firstElementChild;
      hex = entry?.getAttribute('lastClr') || entry?.getAttribute('val') || defaults[color.theme];
    }
    if (color.indexed !== undefined)
      hex =
        color.indexed === 64
          ? '000000'
          : indexed[color.indexed < 8 ? color.indexed : color.indexed - 8];
    if (!hex || !/^[\da-f]{6}$/i.test(hex)) return undefined;
    const tint = (color as Partial<ExcelJS.Color> & { tint?: number }).tint || 0;
    if (tint)
      hex = hex
        .match(/../g)!
        .map((c) => {
          const v = parseInt(c, 16);
          return Math.round(tint < 0 ? v * (1 + tint) : v + (255 - v) * tint)
            .toString(16)
            .padStart(2, '0');
        })
        .join('');
    return '#' + hex;
  };
}
export async function importWorkbook(data: ArrayBuffer): Promise<WorkbookContent> {
  const ExcelJS = (await import('exceljs')).default;
  const zip = await JSZip.loadAsync(data),
    mapping = await workbookSheets(zip);
  const themeText = await zip.file('xl/theme/theme1.xml')?.async('string');
  const color = colorResolver(themeText ? parseXML(themeText) : null);
  const book = new ExcelJS.Workbook();
  let readerZip: JSZip | undefined;
  for (const entry of Object.values(zip.files)) {
    if (!/^xl\/tables\/[^/]+\.xml$/.test(entry.name)) continue;
    const table = parseXML(await entry.async('string'));
    const filters = elements(table, 'autoFilter');
    if (filters.length) {
      // ExcelJS's table parser also rejects valid icon/color filters. Adapt only its reading copy;
      // the original archive used for storage and export is never changed.
      readerZip ||= await JSZip.loadAsync(data);
      filters.forEach((node) => node.remove());
      readerZip.file(entry.name, new XMLSerializer().serializeToString(table), {
        createFolders: false,
      });
    }
  }
  // These features remain in the retained package. ExcelJS cannot parse some valid icon filters;
  // Filter visibility/metadata is read from original XML; conditional-format evaluation remains partial.
  await book.xlsx.load(readerZip ? await readerZip.generateAsync({ type: 'arraybuffer' }) : data, {
    // Read hyperlinks from source XML separately: ExcelJS turns numeric/formula cells with links into text objects.
    // Table relationships/names are read below; ExcelJS assumes tableN.xml and relative targets.
    ignoreNodes: [
      'autoFilter',
      'conditionalFormatting',
      'hyperlinks',
      'dataValidations',
      'tableParts',
    ],
  });
  const sheets: Sheet[] = [];
  for (const ws of book.worksheets) {
    if (ws.rowCount > 10000 || ws.columnCount > 256)
      throw new Error(
        'This workbook exceeds the current editable grid limit of 10,000 rows or 256 columns.',
      );
    const sourcePath = mapping.sheets.find((s) => s.name === ws.name)!.path;
    const xml = parseXML(await zip.file(sourcePath)!.async('string'));
    const autoFilters = readAutoFilters(xml.documentElement, sourcePath);
    const extendedFormulas = elements(xml, 'f').filter((f) =>
      ['array', 'dataTable'].includes(f.getAttribute('t') || ''),
    );
    const tables: NonNullable<Sheet['tables']> = [];
    const relPath = sourcePath.replace(/([^/]+)$/, '_rels/$1.rels');
    const relText = await zip.file(relPath)?.async('string');
    const tableIds = new Set(elements(xml, 'tablePart').map((e) => e.getAttribute('r:id')));
    if (relText)
      for (const rel of elements(parseXML(relText), 'Relationship')) {
        if (
          !tableIds.has(rel.getAttribute('Id')) ||
          rel.getAttribute('TargetMode') === 'External' ||
          !rel.getAttribute('Type')?.endsWith('/table')
        )
          continue;
        const path = resolvePart(sourcePath, rel.getAttribute('Target')!);
        const text = await zip.file(path)?.async('string');
        if (!text) throw new Error('Missing workbook table part.');
        const table = parseXML(text).documentElement;
        const style = elements(table, 'tableStyleInfo')[0];
        autoFilters.push(...readAutoFilters(table, path));
        tables.push({
          sourcePath: path,
          name: table.getAttribute('displayName') || table.getAttribute('name')!,
          ref: table.getAttribute('ref')!,
          columns: elements(table, 'tableColumn').map((e) => e.getAttribute('name') || ''),
          calculatedColumns: elements(table, 'tableColumn').map(
            (e) => elements(e, 'calculatedColumnFormula')[0]?.textContent || null,
          ),
          totalLabels: elements(table, 'tableColumn').map((e) => e.getAttribute('totalsRowLabel')),
          ...(elements(table, 'totalsRowFormula').some((e) => e.attributes.length > 0) ||
          elements(table, 'tableColumn').some(
            (e) => !['', 'none', 'custom'].includes(e.getAttribute('totalsRowFunction') || ''),
          )
            ? {
                totalsActivationBlocked:
                  'Automatic totals activation with preset or extended totals formulas is not supported yet.',
              }
            : {}),
          totalFormulas: elements(table, 'tableColumn').map(
            (e) => elements(e, 'totalsRowFormula')[0]?.textContent || null,
          ),
          ...((table.getAttribute('tableType') &&
            table.getAttribute('tableType') !== 'worksheet') ||
          elements(table, 'extLst').length ||
          elements(table, 'xmlColumnPr').length ||
          elements(table, 'calculatedColumnFormula').some((f) => f.attributes.length > 0) ||
          extendedFormulas.some((f) => {
            try {
              const a = tableBounds(
                  f.getAttribute('ref') || f.parentElement?.getAttribute('r') || '',
                ),
                b = tableBounds(table.getAttribute('ref') || '');
              return a.r <= b.rr && a.rr >= b.r && a.c <= b.cc && a.cc >= b.c;
            } catch {
              return true;
            }
          }) ||
          elements(table, 'tableColumn').some((c) => c.hasAttribute('queryTableFieldId'))
            ? {
                calculationBlocked:
                  'Automatic calculated columns are unavailable for this table’s external data or extended formula metadata.',
              }
            : {}),
          headerRows: Number(table.getAttribute('headerRowCount') ?? 1),
          totalRows: Number(table.getAttribute('totalsRowCount') ?? 0),
          ...(style
            ? {
                style: {
                  name: style.getAttribute('name') || '',
                  rowStripes: ['1', 'true'].includes(style.getAttribute('showRowStripes') || ''),
                  columnStripes: ['1', 'true'].includes(
                    style.getAttribute('showColumnStripes') || '',
                  ),
                  firstColumn: ['1', 'true'].includes(style.getAttribute('showFirstColumn') || ''),
                  lastColumn: ['1', 'true'].includes(style.getAttribute('showLastColumn') || ''),
                },
              }
            : {}),
          ...((table.getAttribute('tableType') &&
            table.getAttribute('tableType') !== 'worksheet') ||
          extendedFormulas.length ||
          elements(table, 'extLst').length ||
          elements(table, 'sortState').length ||
          elements(table, 'xmlColumnPr').length ||
          elements(table, 'tableColumn').some((c) => c.hasAttribute('queryTableFieldId')) ||
          Object.keys(zip.files).some((p) =>
            /^xl\/(pivotTables|pivotCache|slicers|slicerCaches)\//.test(p),
          )
            ? {
                resizeBlocked:
                  'Resizing tables with array formulas, sorting, external data, pivots or slicers is not supported yet.',
              }
            : {}),
        });
      }
    const originalCells = new Map(elements(xml, 'c').map((e) => [e.getAttribute('r'), e]));
    const validationRanges = readValidationRanges(xml.documentElement);
    const cells: Record<string, Cell> = {};
    const rowHeights: Record<string, number> = {},
      columnWidths: Record<string, number> = {};
    const hiddenRows: number[] = [],
      hiddenColumns: number[] = [];
    ws.eachRow({ includeEmpty: true }, (row) => {
      if (row.height) rowHeights[row.number - 1] = (row.height * 4) / 3;
      if (row.hidden) hiddenRows.push(row.number - 1);
      row.eachCell({ includeEmpty: true }, (cell) => {
        if (!originalCells.has(cell.address)) return;
        const raw = cell.value,
          xmlCell = originalCells.get(cell.address)!;
        const rawValue = elements(xmlCell, 'v')[0]?.textContent;
        let value = '',
          dataType: Cell['dataType'];
        if (cell.formula) value = '=' + cell.formula;
        else if (raw instanceof Date) {
          value = rawValue || '';
          dataType = 'date';
        } else if (typeof raw === 'string') {
          value = raw;
          dataType = 'text';
        } else if (typeof raw === 'number') {
          value = String(raw);
          dataType = 'number';
        } else if (typeof raw === 'boolean') {
          value = raw ? 'TRUE' : 'FALSE';
          dataType = 'boolean';
        } else if (raw && typeof raw === 'object') {
          value =
            'richText' in raw
              ? raw.richText.map((t) => t.text).join('')
              : 'text' in raw
                ? raw.text
                : 'error' in raw
                  ? raw.error
                  : cell.text;
          dataType = 'error' in raw ? 'error' : 'text';
        }
        const border = (side: 'top' | 'bottom' | 'left' | 'right') => {
          const b = cell.border?.[side];
          if (!b?.style) return undefined;
          return `${b.style === 'thick' ? 3 : b.style === 'medium' ? 2 : 1}px ${b.style === 'double' ? 'double' : /dash/i.test(b.style) ? 'dashed' : b.style === 'dotted' ? 'dotted' : 'solid'} ${color(b.color) || '#000000'}`;
        };
        const result = cell.result;
        const validation = cellValidation({ cells: {}, validationRanges } as Sheet, cell.address);
        const note =
          typeof cell.note === 'string' ? cell.note : cell.note?.texts?.map((t) => t.text).join('');
        cells[cell.address] = {
          value,
          dataType,
          cachedValue:
            result instanceof Date
              ? Number(rawValue)
              : typeof result === 'number' ||
                  typeof result === 'string' ||
                  typeof result === 'boolean'
                ? result
                : undefined,
          bold: cell.font?.bold,
          italic: cell.font?.italic,
          underline: !!cell.font?.underline,
          fontFamily: cell.font?.name,
          fontSize: cell.font?.size,
          color: color(cell.font?.color),
          fill: cell.fill?.type === 'pattern' ? color(cell.fill.fgColor) : undefined,
          numFmt: cell.numFmt || undefined,
          date1904: book.properties.date1904 || false,
          align: ['left', 'center', 'right'].includes(cell.alignment?.horizontal || '')
            ? (cell.alignment.horizontal as Cell['align'])
            : undefined,
          vertical: ['top', 'middle', 'bottom'].includes(cell.alignment?.vertical || '')
            ? (cell.alignment.vertical as Cell['vertical'])
            : undefined,
          wrap: cell.alignment?.wrapText,
          locked: cell.protection?.locked !== false,
          borders: {
            top: border('top'),
            bottom: border('bottom'),
            left: border('left'),
            right: border('right'),
          },
          hyperlink: cell.hyperlink || undefined,
          validation,
          note,
        };
        // Merged followers carry geometry/borders in the package, not duplicate editable values.
        if (cell.isMerged && cell.master.address !== cell.address) cells[cell.address].value = '';
      });
    });
    for (let c = 1; c <= Math.max(26, ws.columnCount); c++) {
      const col = ws.getColumn(c);
      if (col.width !== undefined) columnWidths[c - 1] = Math.max(4, Math.floor(col.width * 7 + 5));
      if (col.hidden) hiddenColumns.push(c - 1);
    }
    const view = ws.views?.[0];
    const defaultRowHeight = ((ws.properties.defaultRowHeight || 15) * 4) / 3;
    const defaultColumnWidth = (ws.properties.defaultColWidth || 8.43) * 7 + 5;
    const definedNames: Record<string, string> = {};
    const sheetIndex = mapping.sheets.findIndex((s) => s.name === ws.name);
    const nameDefinitions: NonNullable<Sheet['nameDefinitions']> = elements(
      mapping.book,
      'definedName',
    )
      .filter(
        (e) =>
          !e.hasAttribute('localSheetId') || e.getAttribute('localSheetId') === String(sheetIndex),
      )
      .map((e) => ({
        name: e.getAttribute('name') || '',
        formula: e.textContent || '',
        scope: e.hasAttribute('localSheetId') ? 'worksheet' : 'workbook',
        referenceOrigin: 'ooxml-a1',
      }));
    for (const entry of elements(mapping.book, 'definedName').filter(
      (e) => !e.hasAttribute('localSheetId'),
    ))
      definedNames[entry.getAttribute('name')!.toLowerCase()] = entry.textContent || '';
    for (const entry of elements(mapping.book, 'definedName').filter(
      (e) => e.getAttribute('localSheetId') === String(sheetIndex),
    ))
      definedNames[entry.getAttribute('name')!.toLowerCase()] = entry.textContent || '';
    const images: NonNullable<Sheet['images']> = [];
    const position = (anchor: ExcelJS.Anchor) => ({
      x:
        Array.from({ length: anchor.nativeCol }, (_, c) =>
          hiddenColumns.includes(c) ? 0 : columnWidths[c] || defaultColumnWidth,
        ).reduce((a, b) => a + b, 0) +
        anchor.nativeColOff / 9525,
      y:
        Array.from({ length: anchor.nativeRow }, (_, r) =>
          hiddenRows.includes(r) ? 0 : rowHeights[r] || defaultRowHeight,
        ).reduce((a, b) => a + b, 0) +
        anchor.nativeRowOff / 9525,
    });
    for (const placement of ws.getImages()) {
      const image = book.getImage(Number(placement.imageId));
      if (!image || !['png', 'jpeg', 'gif'].includes(image.extension)) continue;
      let src = image.base64;
      if (image.buffer) {
        const bytes = new Uint8Array(image.buffer);
        let binary = '';
        for (let i = 0; i < bytes.length; i += 8192)
          binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
        src = `data:image/${image.extension};base64,${btoa(binary)}`;
      }
      if (!src) continue;
      if (!src.startsWith('data:')) src = `data:image/${image.extension};base64,${src}`;
      const start = position(placement.range.tl),
        ext = (placement.range as ExcelJS.ImageRange & { ext?: { width: number; height: number } })
          .ext;
      const end = placement.range.br ? position(placement.range.br) : null;
      images.push({
        src,
        ...start,
        w: Math.max(1, ext?.width || (end ? end.x - start.x : 100)),
        h: Math.max(1, ext?.height || (end ? end.y - start.y : 100)),
      });
    }
    readWorksheetLinks(xml.documentElement, relText, cells);
    const arrayFormulas = extendedFormulas
      .filter((f) => f.getAttribute('t') === 'array')
      .map((f) => ({
        anchor: f.parentElement!.getAttribute('r')!,
        ref: (f.getAttribute('ref') || f.parentElement!.getAttribute('r')!)
          .replaceAll('$', '')
          .toUpperCase(),
      }));
    const anchors = new Set<string>();
    for (const array of arrayFormulas) {
      if (
        anchors.has(array.anchor) ||
        array.ref.split(':')[0] !== array.anchor ||
        !arrayContains(array.ref, array.anchor) ||
        !cells[array.anchor]?.value.startsWith('=')
      )
        throw Error('Invalid array formula source identity.');
      anchors.add(array.anchor);
    }
    sheets.push({
      id: uid(),
      name: ws.name,
      sourcePath,
      date1904: book.properties.date1904 || false,
      cells,
      arrayFormulas,
      rows: Object.keys(cells).reduce(
        (max, ref) => Math.max(max, coordinates(ref)[0] + 1),
        Math.max(100, ws.rowCount),
      ),
      cols: Object.keys(cells).reduce(
        (max, ref) => Math.max(max, coordinates(ref)[1] + 1),
        Math.max(26, ws.columnCount),
      ),
      state: ws.state || 'visible',
      rowHeights,
      columnWidths,
      defaultRowHeight,
      defaultColumnWidth,
      hiddenRows,
      ...(elements(xml, 'sheetPr')[0]?.getAttribute('filterMode') === '1'
        ? { filterMode: true }
        : {}),
      ...(autoFilters.length ? { autoFilters } : {}),
      hiddenColumns,
      images,
      definedNames,
      nameDefinitions,
      tables,
      merges: elements(xml, 'mergeCell').map((e) => e.getAttribute('ref')!),
      validationRanges,
      protected: elements(xml, 'sheetProtection').length > 0,
      tableTheme: Array.from({ length: 6 }, (_, i) => color({ theme: i + 4 })!),
      ...(elements(mapping.book, 'workbookProtection').length
        ? { tableEditingBlocked: 'This workbook is protected.' }
        : {}),
      gridLines: view?.showGridLines !== false,
      frozenRows: view?.state === 'frozen' ? view.ySplit || 0 : 0,
      frozenColumns: view?.state === 'frozen' ? view.xSplit || 0 : 0,
    });
  }
  if (!sheets.length) throw new Error('This workbook has no worksheets.');
  return { kind: 'excel', sheets };
}
