import type JSZip from 'jszip';
import type { Sheet } from './model';
import { elements, parseXML, spreadsheetNS as NS } from './xlsx-import';
import { editSheetTable, tableBounds, tableHeaderError, type SheetTable } from './sheet-tables';

const relationshipNS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const serialize = (doc: Document) => new XMLSerializer().serializeToString(doc);
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
function calculatedNodes(root: Element, before: SheetTable | undefined, after: SheetTable) {
  if (same(before?.calculatedColumns, after.calculatedColumns)) return;
  if (after.calculationBlocked) throw Error(after.calculationBlocked);
  if (!after.calculatedColumns || after.calculatedColumns.length !== after.columns.length)
    throw Error('Calculated-column metadata must match the table columns.');
  const columns = elements(root, 'tableColumn');
  for (const [i, formula] of after.calculatedColumns.entries()) {
    if ((before?.calculatedColumns?.[i] ?? null) === formula) continue;
    if (
      formula !== null &&
      (!formula ||
        formula.length > 8192 ||
        formula.startsWith('=') ||
        /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(formula))
    )
      throw Error('Invalid calculated-column formula.');
    const column = columns[i];
    if (!column) throw Error('The calculated column is missing from its source table.');
    let node = elements(column, 'calculatedColumnFormula')[0];
    if (node?.attributes.length)
      throw Error('Editing extended calculated-column formulas is not supported.');
    if (formula === null) {
      node?.remove();
      continue;
    }
    if (!node) {
      node = root.ownerDocument.createElementNS(NS, 'calculatedColumnFormula');
      column.insertBefore(node, column.firstChild);
    }
    node.textContent = formula;
  }
}
function styleNode(root: Element, table: SheetTable) {
  if (!table.style) return;
  let style = elements(root, 'tableStyleInfo')[0];
  if (!style) {
    style = root.ownerDocument.createElementNS(NS, 'tableStyleInfo');
    root.insertBefore(style, elements(root, 'extLst')[0] || null);
  }
  style.setAttribute('name', table.style.name);
  for (const [key, attr] of [
    ['rowStripes', 'showRowStripes'],
    ['columnStripes', 'showColumnStripes'],
    ['firstColumn', 'showFirstColumn'],
    ['lastColumn', 'showLastColumn'],
  ] as const)
    style.setAttribute(attr, table.style[key] ? '1' : '0');
}

function totalNodes(root: Element, before: SheetTable | undefined, after: SheetTable) {
  if ([after.totalFormulas, after.totalLabels].some((a) => a && a.length !== after.columns.length))
    throw Error('Totals metadata must match the table columns.');
  if (before?.totalRows !== after.totalRows) {
    root.setAttribute('totalsRowCount', String(after.totalRows));
    root.setAttribute('totalsRowShown', after.totalRows ? '1' : '0');
  }
  if (
    same(before?.totalFormulas, after.totalFormulas) &&
    same(before?.totalLabels, after.totalLabels)
  )
    return;
  const columns = elements(root, 'tableColumn');
  for (let i = 0; i < after.columns.length; i++) {
    const formula = after.totalFormulas?.[i] ?? null,
      label = after.totalLabels?.[i] ?? null;
    if (
      (formula !== null && !formula) ||
      (formula && label) ||
      (label && (label.length > 32767 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(label)))
    )
      throw Error('Invalid totals metadata.');
    if (
      formula === (before?.totalFormulas?.[i] ?? null) &&
      label === (before?.totalLabels?.[i] ?? null)
    )
      continue;
    const column = columns[i];
    if (!column) throw Error('Missing totals column.');
    let node = elements(column, 'totalsRowFormula')[0];
    if (node?.attributes.length) throw Error('Extended totals formulas cannot be edited.');
    if (formula) {
      if (
        formula.length > 8192 ||
        formula.startsWith('=') ||
        /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(formula)
      )
        throw Error('Invalid totals formula.');
      if (!node) {
        node = root.ownerDocument.createElementNS(NS, 'totalsRowFormula');
        column.insertBefore(
          node,
          elements(column, 'xmlColumnPr')[0] || elements(column, 'extLst')[0] || null,
        );
      }
      node.textContent = formula;
      column.setAttribute('totalsRowFunction', 'custom');
      column.removeAttribute('totalsRowLabel');
    } else {
      node?.remove();
      column.removeAttribute('totalsRowFunction');
      if (label) column.setAttribute('totalsRowLabel', label);
      else column.removeAttribute('totalsRowLabel');
    }
  }
}

/** Patch table definitions independently of worksheet cells. No table is rebuilt on style/row-range edits. */
export async function writeSheetTables(
  zip: JSZip,
  sheetPath: string,
  root: Element,
  before: Sheet,
  after: Sheet,
): Promise<boolean> {
  const headerError = tableHeaderError(before, after);
  if (headerError) throw Error(headerError);
  const previous = before.tables || [],
    next = after.tables || [];
  if (same(previous, next)) return false;
  if (after.protected || after.tableEditingBlocked)
    throw Error(after.tableEditingBlocked || 'This worksheet is protected.');
  if (previous.some((t) => !next.some((n) => n.sourcePath === t.sourcePath && n.name === t.name)))
    throw Error('Removing or renaming existing tables requires reference repair.');
  let sheetDirty = false;
  for (const table of next) {
    const old = previous.find((t) => t.sourcePath === table.sourcePath);
    if (same(old, table)) continue;
    if (old) {
      if (
        !same(
          {
            ...old,
            ref: table.ref,
            style: table.style,
            calculatedColumns: table.calculatedColumns,
            totalRows: table.totalRows,
            totalFormulas: table.totalFormulas,
            totalLabels: table.totalLabels,
          },
          table,
        )
      )
        throw Error('Editing table column or source metadata is not supported.');
      if (
        old.totalRows !== table.totalRows &&
        !(
          old.totalRows === 0 &&
          table.totalRows === 1 &&
          tableBounds(table.ref).rr > tableBounds(old.ref).rr
        )
      )
        throw Error('Unsupported totals-row structure change.');
      if (old.totalRows !== table.totalRows && old.totalsActivationBlocked)
        throw Error(old.totalsActivationBlocked);
      editSheetTable(
        { kind: 'excel', sheets: [before] },
        before.id,
        old.name,
        table.ref,
        table.style || {
          name: '',
          rowStripes: false,
          columnStripes: false,
          firstColumn: false,
          lastColumn: false,
        },
        true,
      );
      if (!old.sourcePath || !zip.file(old.sourcePath)) throw Error('Missing original table part.');
      const doc = parseXML(await zip.file(old.sourcePath)!.async('string')),
        tableRoot = doc.documentElement;
      tableRoot.setAttribute('ref', table.ref);
      if (old.ref !== table.ref)
        for (const filter of elements(tableRoot, 'autoFilter'))
          filter.setAttribute('ref', table.ref);
      if (!same(old.style, table.style)) styleNode(tableRoot, table);
      calculatedNodes(tableRoot, old, table);
      totalNodes(tableRoot, old, table);
      zip.file(old.sourcePath, serialize(doc));
      continue;
    }
    const path = table.sourcePath;
    if (!path || !/^xl\/tables\/noffice[a-zA-Z0-9_-]+\.xml$/.test(path) || zip.file(path))
      throw Error('The new table part conflicts with an existing file.');
    const b = tableBounds(table.ref);
    if (
      table.headerRows !== 1 ||
      table.columns.length !== b.cc - b.c + 1 ||
      b.rr <= b.r + table.totalRows
    )
      throw Error('Invalid new table definition.');
    // Allocate a package-wide numeric table ID; names and part paths remain independent.
    let id = 1;
    for (const entry of Object.keys(zip.files).filter((p) => /^xl\/tables\/.*\.xml$/.test(p))) {
      const existing = Number(
        parseXML(await zip.file(entry)!.async('string')).documentElement.getAttribute('id'),
      );
      if (Number.isInteger(existing)) id = Math.max(id, existing + 1);
    }
    const doc = parseXML(`<table xmlns="${NS}"/>`),
      t = doc.documentElement;
    for (const [k, v] of Object.entries({
      id,
      name: table.name,
      displayName: table.name,
      ref: table.ref,
      headerRowCount: 1,
      totalsRowShown: 0,
    }))
      t.setAttribute(k, String(v));
    const filter = doc.createElementNS(NS, 'autoFilter');
    filter.setAttribute('ref', table.ref);
    t.appendChild(filter);
    const columns = doc.createElementNS(NS, 'tableColumns');
    columns.setAttribute('count', String(table.columns.length));
    t.appendChild(columns);
    table.columns.forEach((name, i) => {
      const c = doc.createElementNS(NS, 'tableColumn');
      c.setAttribute('id', String(i + 1));
      c.setAttribute('name', name);
      columns.appendChild(c);
    });
    calculatedNodes(t, undefined, table);
    totalNodes(t, undefined, table);
    styleNode(t, table);
    zip.file(path, serialize(doc));
    const relPath = sheetPath.replace(/([^/]+)$/, '_rels/$1.rels');
    const relText = await zip.file(relPath)?.async('string'),
      relDoc = parseXML(relText || `<Relationships xmlns="${relationshipNS}"/>`);
    const used = new Set(elements(relDoc, 'Relationship').map((r) => r.getAttribute('Id')));
    let n = 1;
    while (used.has(`rId${n}`)) n++;
    const rel = relDoc.createElementNS(relationshipNS, 'Relationship');
    rel.setAttribute('Id', `rId${n}`);
    rel.setAttribute('Type', officeRel + '/table');
    rel.setAttribute('Target', '/' + path);
    relDoc.documentElement.appendChild(rel);
    zip.file(relPath, serialize(relDoc));
    let parts = elements(root, 'tableParts')[0];
    if (!parts) {
      parts = root.ownerDocument.createElementNS(NS, 'tableParts');
      root.insertBefore(
        parts,
        Array.from(root.children).find((e) => ['extLst'].includes(e.localName)) || null,
      );
    }
    const part = root.ownerDocument.createElementNS(NS, 'tablePart');
    part.setAttributeNS(officeRel, 'r:id', `rId${n}`);
    parts.appendChild(part);
    parts.setAttribute('count', String(parts.children.length));
    const types = parseXML(await zip.file('[Content_Types].xml')!.async('string'));
    const override = types.createElementNS(types.documentElement.namespaceURI, 'Override');
    override.setAttribute('PartName', '/' + path);
    override.setAttribute(
      'ContentType',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
    );
    types.documentElement.appendChild(override);
    zip.file('[Content_Types].xml', serialize(types));
    sheetDirty = true;
  }
  return sheetDirty;
}
