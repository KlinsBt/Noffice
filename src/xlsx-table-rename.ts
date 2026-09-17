import JSZip from 'jszip';
import { formulaCellIdentity, staticFormulaDependents } from './formula-static-dependencies';
import { address, coordinates, calculator } from './formulas';
import {
  elements,
  parseXML,
  resolvePart,
  spreadsheetNS as NS,
  workbookSheets,
  importWorkbook,
} from './xlsx-import';
import { tableBounds } from './sheet-tables';
import { renameTableReferences, type ReferenceRename } from './table-references';

const serialize = (doc: Document) => new XMLSerializer().serializeToString(doc);
export interface PackageTableRename extends ReferenceRename {
  sourcePath: string;
  columnIndex?: number;
}

/** Rename a table in a copy of its package, including dependencies outside the editable grid. */
export async function renameXlsxTable(
  data: ArrayBuffer,
  edit: PackageTableRename,
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(data),
    mapping = await workbookSheets(zip);
  if (
    Object.keys(zip.files).some((p) =>
      /^xl\/(?:pivotTables|pivotCache|slicers|slicerCaches|queryTables|externalLinks|model)\/|^xl\/(?:connections\.xml|vbaProject\.bin)/i.test(
        p,
      ),
    )
  )
    throw Error(
      'Table renaming with pivots, slicers, queries, external links, data models or macros is not supported yet.',
    );
  if (elements(mapping.book, 'workbookProtection').length)
    throw Error('This workbook is protected.');
  const docs = new Map<string, Document>([['xl/workbook.xml', mapping.book]]);
  const ownership = new Map<string, string>();
  for (const s of mapping.sheets) {
    if (!/\/worksheets\//.test(s.path))
      throw Error('Only standard worksheets support table renaming.');
    const doc = parseXML(await zip.file(s.path)!.async('string'));
    docs.set(s.path, doc);
    const rel = zip.file(s.path.replace(/([^/]+)$/, '_rels/$1.rels'));
    if (rel)
      for (const r of elements(parseXML(await rel.async('string')), 'Relationship'))
        if (
          r.getAttribute('Type')?.endsWith('/table') &&
          r.getAttribute('TargetMode') !== 'External'
        )
          ownership.set(resolvePart(s.path, r.getAttribute('Target')!), s.path);
  }
  for (const path of Object.keys(zip.files).filter((p) =>
    /^xl\/(?:tables|charts)\/[^/]+\.xml$/.test(p),
  ))
    docs.set(path, parseXML(await zip.file(path)!.async('string')));
  const table = docs.get(edit.sourcePath)?.documentElement;
  if (!table || table.localName !== 'table' || table.getAttribute('displayName') !== edit.table)
    throw Error('The original table definition no longer matches this workbook.');
  const owner = ownership.get(edit.sourcePath),
    sheet = owner && docs.get(owner);
  if (!sheet) throw Error('The table worksheet relationship is missing.');
  if (elements(sheet, 'sheetProtection').length) throw Error('This worksheet is protected.');
  for (const doc of docs.values()) {
    if (elements(doc, 'extLst').length || elements(doc, 'AlternateContent').length)
      throw Error('Extension references cannot yet be repaired during table renaming.');
    if (
      elements(doc, 'table').some(
        (t) => t.getAttribute('tableType') && t.getAttribute('tableType') !== 'worksheet',
      )
    )
      throw Error('External-data tables cannot yet be renamed.');
  }
  const b = tableBounds(table.getAttribute('ref')!);
  const columns = elements(table, 'tableColumn');
  const columnEdits = (edit.columns || (edit.column ? [edit.column] : [])).map((c) => ({
    ...c,
    index: columns.findIndex((node) => node.getAttribute('name') === c.before),
  }));
  if (
    columnEdits.some((c) => c.index < 0) ||
    new Set(columnEdits.map((c) => c.index)).size !== columnEdits.length
  )
    throw Error('The original column definitions no longer match this workbook.');
  if (
    edit.column &&
    (!Number.isInteger(edit.columnIndex) ||
      edit.columnIndex! < 0 ||
      columns[edit.columnIndex!]?.getAttribute('name') !== edit.column.before)
  )
    throw Error('The original column definition no longer matches this workbook.');
  const original = new Map([...docs].map(([p, d]) => [p, serialize(d)]));
  const tables = [...docs]
    .filter(([, d]) => d.documentElement.localName === 'table')
    .map(([path, d]) => ({
      path,
      sheet: ownership.get(path),
      name: d.documentElement.getAttribute('displayName')!,
      bounds: tableBounds(d.documentElement.getAttribute('ref')!),
    }));
  for (const [path, doc] of docs) {
    const context = tables.find((t) => t.path === path)?.name;
    for (const f of Array.from(doc.getElementsByTagName('*')).filter((e) =>
      [
        'f',
        'formula',
        'formula1',
        'formula2',
        'definedName',
        'calculatedColumnFormula',
        'totalsRowFormula',
      ].includes(e.localName),
    )) {
      let local = context;
      if (f.localName === 'f' && f.parentElement?.localName === 'c') {
        const ref = f.parentElement.getAttribute('r');
        if (ref) {
          const [r, c] = coordinates(ref);
          local = tables.find(
            (t) =>
              t.sheet === path &&
              r >= t.bounds.r &&
              r <= t.bounds.rr &&
              c >= t.bounds.c &&
              c <= t.bounds.cc,
          )?.name;
        }
      }
      const old = f.textContent || '',
        next = renameTableReferences(old, edit, local);
      if (next !== old) {
        if (/(?:^|[^\w.])(?:_xlfn\.)?(?:LET|LAMBDA)\s*\(/i.test(old))
          throw Error('Scoped LET/LAMBDA references cannot yet be repaired during table renaming.');
        f.textContent = next;
      }
    }
  }
  if (edit.name !== edit.table) {
    table.setAttribute('name', edit.name);
    table.setAttribute('displayName', edit.name);
  }
  for (const change of columnEdits) {
    columns[change.index].setAttribute('name', change.after);
    if (table.getAttribute('headerRowCount') !== '0') {
      const ref = address(b.r, b.c + change.index);
      const cell = elements(sheet, 'c').find((c) => c.getAttribute('r') === ref);
      if (!cell) throw Error('The table header cell is missing.');
      for (const child of Array.from(cell.children))
        if (['f', 'v', 'is'].includes(child.localName)) child.remove();
      cell.setAttribute('t', 'inlineStr');
      const value = sheet.createElementNS(NS, 'is'),
        text = sheet.createElementNS(NS, 't');
      text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      text.textContent = change.after;
      value.appendChild(text);
      cell.insertBefore(value, cell.firstChild);
    }
  }
  // Excel also recalculates unsupported expressions and literal INDIRECT destinations on opening.
  let calc = elements(mapping.book, 'calcPr')[0];
  if (!calc) {
    calc = mapping.book.createElementNS(NS, 'calcPr');
    mapping.book.documentElement.appendChild(calc);
  }
  calc.setAttribute('fullCalcOnLoad', '1');
  calc.setAttribute('forceFullCalc', '1');
  for (const [path, doc] of docs) {
    const text = serialize(doc);
    if (text !== original.get(path)) zip.file(path, text);
  }
  const snapshot = await importWorkbook(await zip.generateAsync({ type: 'arraybuffer' }));
  const calcValues = calculator(snapshot.sheets);
  const changed = new Set<string>();
  const ownerSheet = snapshot.sheets.find((s) => s.sourcePath === owner)!;
  for (const change of columnEdits)
    changed.add(formulaCellIdentity(ownerSheet.id, address(b.r, b.c + change.index)));
  const affected = staticFormulaDependents(snapshot.sheets, changed);
  for (const s of snapshot.sheets) {
    const doc = docs.get(s.sourcePath!)!;
    let dirty = false;
    for (const cell of elements(doc, 'c')) {
      const ref = cell.getAttribute('r')!;
      if (!elements(cell, 'f').length || !affected.has(formulaCellIdentity(s.id, ref))) continue;
      const result = calcValues.result(s, ref);
      if (result.kind === 'unsupported') continue;
      const value =
        typeof result.value === 'boolean' ? (result.value ? '1' : '0') : String(result.value);
      const type =
        typeof result.value === 'string'
          ? result.kind === 'error'
            ? 'e'
            : 'str'
          : typeof result.value === 'boolean'
            ? 'b'
            : null;
      let v = elements(cell, 'v')[0];
      if (v?.textContent === value && cell.getAttribute('t') === type) continue;
      if (!v) {
        v = doc.createElementNS(NS, 'v');
        cell.appendChild(v);
      }
      v.textContent = value;
      if (type) cell.setAttribute('t', type);
      else cell.removeAttribute('t');
      dirty = true;
    }
    if (dirty) zip.file(s.sourcePath!, serialize(doc));
  }
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
