import JSZip from 'jszip';
import { formulaCellIdentity, staticFormulaDependents } from './formula-static-dependencies';
import type { Cell, OfficeFile, Sheet, WorkbookContent } from './model';
import { coordinates, calculator } from './formulas';
import { writeFilterColumns } from './sheet-filters';
import { writeFrozenPanes } from './sheet-panes';
import { writeWorksheetLinks } from './xlsx-links';
import { writeValidationChanges } from './xlsx-validation';
import { writeSheetTables } from './xlsx-tables';
import { restoreTableMetadata } from './sheet-tables';
import { arrayEditError, singleArray } from './sheet-arrays';
import {
  elements,
  importWorkbook,
  parseXML,
  resolvePart,
  spreadsheetNS as NS,
  workbookSheets,
} from './xlsx-import';

const equal = (a: unknown, b: unknown): boolean => {
  if (a === b) return true;
  if (!a || !b || typeof a !== 'object' || typeof b !== 'object') return false;
  const aa = a as Record<string, unknown>,
    bb = b as Record<string, unknown>;
  const keys = new Set([...Object.keys(aa), ...Object.keys(bb)]);
  return [...keys].every((k) => equal(aa[k], bb[k]));
};
const serialize = (doc: Document) => new XMLSerializer().serializeToString(doc);
function child(parent: Element, name: string) {
  return Array.from(parent.children).find((e) => e.localName === name);
}
function ensure(parent: Element, name: string) {
  let e = child(parent, name);
  if (!e) {
    e = parent.ownerDocument.createElementNS(parent.namespaceURI || NS, name);
    const order =
      parent.localName === 'c'
        ? ['f', 'v', 'is', 'extLst']
        : parent.localName === 'workbook'
          ? [
              'fileVersion',
              'fileSharing',
              'workbookPr',
              'workbookProtection',
              'bookViews',
              'sheets',
              'functionGroups',
              'externalReferences',
              'definedNames',
              'calcPr',
              'oleSize',
              'customWorkbookViews',
              'pivotCaches',
              'smartTagPr',
              'smartTagTypes',
              'webPublishing',
              'fileRecoveryPr',
              'webPublishObjects',
              'extLst',
            ]
          : [];
    const next = Array.from(parent.children).find(
      (c) => order.indexOf(c.localName) > order.indexOf(name),
    );
    parent.insertBefore(e, next || null);
  }
  return e;
}
function remove(parent: Element, name: string) {
  child(parent, name)?.remove();
}

/** Patch only changed cells/properties. Untouched ZIP entries keep their uncompressed bytes. */
export async function exportRetainedWorkbook(file: OfficeFile): Promise<Blob> {
  let current = file.content as WorkbookContent;
  const original = current.xlsxStructureBase
    ? Uint8Array.from(atob(current.xlsxStructureBase), (c) => c.charCodeAt(0)).buffer
    : file.original!.data;
  const baseline = await importWorkbook(original);
  current = restoreTableMetadata(current, baseline);
  for (const sheet of current.sheets) {
    const old = baseline.sheets.find((s) => s.sourcePath === sheet.sourcePath);
    if (old) {
      const error = arrayEditError(old, sheet);
      if (error) throw Error(error);
    }
  }
  const zip = await JSZip.loadAsync(original),
    mapping = await workbookSheets(zip);
  // Structural edits require relationship/reference repair, never a destructive fallback rebuild.
  if (
    current.sheets.length !== baseline.sheets.length ||
    current.sheets.some(
      (s, i) =>
        s.sourcePath !== baseline.sheets[i].sourcePath || s.name !== baseline.sheets[i].name,
    )
  )
    throw new Error(
      'Imported workbook sheet structure changed. Export requires its original sheet names and order to preserve references and relationships.',
    );
  for (let i = 0; i < current.sheets.length; i++)
    for (const key of [
      'merges',
      'hiddenColumns',
      'images',
      'protected',
      'definedNames',
      'nameDefinitions',
      'date1904',
      'gridLines',
    ] as const) {
      if (!equal(current.sheets[i][key], baseline.sheets[i][key]))
        throw new Error(
          `Editing ${key} in an imported workbook is not supported by the preservation writer yet.`,
        );
    }
  if (current.sheets.every((s, i) => equal({ ...s, id: '' }, { ...baseline.sheets[i], id: '' })))
    return new Blob([original], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  const styles = parseXML(await zip.file('xl/styles.xml')!.async('string'));
  const styleRoot = styles.documentElement;
  let stylesChanged = false,
    valuesChanged = current.sheets.some((s, i) => !equal(s.tables, baseline.sheets[i].tables)),
    formulasChanged = false,
    workbookChanged = false;
  let changedCells = new Set<string>();
  const appendStyle = (group: string, node: Element) => {
    const parent = ensure(styleRoot, group),
      source = serializeElement(node);
    const index = Array.from(parent.children).findIndex((e) => serializeElement(e) === source);
    if (index >= 0) return index;
    const id = parent.children.length;
    parent.appendChild(node);
    parent.setAttribute('count', String(parent.children.length));
    stylesChanged = true;
    return id;
  };
  const calc = calculator(current.sheets);
  function styleFor(existing: Element, old: Cell | undefined, next: Cell) {
    const previousId = Number(existing.getAttribute('s') || 0);
    const xf = styles.importNode(
      ensure(styleRoot, 'cellXfs').children[previousId],
      true,
    ) as Element;
    let changed = false;
    const differs = (key: keyof Cell) => key in next && !equal(old?.[key], next[key]);
    const fontKeys = ['fontFamily', 'fontSize', 'bold', 'italic', 'underline', 'color'] as const;
    if (fontKeys.some(differs)) {
      const font = styles.importNode(
        ensure(styleRoot, 'fonts').children[Number(xf.getAttribute('fontId') || 0)],
        true,
      ) as Element;
      for (const key of fontKeys)
        if (differs(key)) {
          const tag = {
            fontFamily: 'name',
            fontSize: 'sz',
            bold: 'b',
            italic: 'i',
            underline: 'u',
            color: 'color',
          }[key];
          remove(font, tag);
          const value = next[key];
          if (value !== undefined && value !== false) {
            const e = ensure(font, tag);
            if (key === 'color') e.setAttribute('rgb', 'FF' + String(value).replace('#', ''));
            else if (key === 'fontFamily' || key === 'fontSize')
              e.setAttribute('val', String(value));
          }
        }
      xf.setAttribute('fontId', String(appendStyle('fonts', font)));
      xf.setAttribute('applyFont', '1');
      changed = true;
    }
    if (differs('fill')) {
      const fill = styles.createElementNS(styleRoot.namespaceURI || NS, 'fill'),
        pattern = ensure(fill, 'patternFill');
      pattern.setAttribute('patternType', next.fill ? 'solid' : 'none');
      if (next.fill)
        ensure(pattern, 'fgColor').setAttribute('rgb', 'FF' + next.fill.replace('#', ''));
      xf.setAttribute('fillId', String(appendStyle('fills', fill)));
      xf.setAttribute('applyFill', '1');
      changed = true;
    }
    if (['align', 'vertical', 'wrap'].some((k) => differs(k as keyof Cell))) {
      const align = ensure(xf, 'alignment');
      for (const [key, attr] of [
        ['align', 'horizontal'],
        ['vertical', 'vertical'],
        ['wrap', 'wrapText'],
      ] as const)
        if (differs(key)) {
          if (next[key] === undefined) align.removeAttribute(attr);
          else
            align.setAttribute(
              attr,
              key === 'wrap'
                ? next.wrap
                  ? '1'
                  : '0'
                : key === 'vertical' && next.vertical === 'middle'
                  ? 'center'
                  : (next[key] as string),
            );
        }
      xf.setAttribute('applyAlignment', '1');
      changed = true;
    }
    if (differs('numFmt') || differs('format')) {
      const code =
        next.numFmt ||
        { general: 'General', number: '0.00', currency: '$#,##0.00', percent: '0.00%' }[
          next.format || 'general'
        ];
      const formats = ensure(styleRoot, 'numFmts');
      // numFmts precedes fonts in the SpreadsheetML stylesheet sequence.
      styleRoot.insertBefore(formats, styleRoot.firstChild);
      let fmt = Array.from(formats.children).find((e) => e.getAttribute('formatCode') === code);
      if (!fmt) {
        fmt = styles.createElementNS(styleRoot.namespaceURI || NS, 'numFmt');
        fmt.setAttribute(
          'numFmtId',
          String(
            Math.max(
              163,
              ...Array.from(formats.children).map((e) => Number(e.getAttribute('numFmtId'))),
            ) + 1,
          ),
        );
        fmt.setAttribute('formatCode', code);
        formats.appendChild(fmt);
        formats.setAttribute('count', String(formats.children.length));
        stylesChanged = true;
      }
      xf.setAttribute('numFmtId', fmt.getAttribute('numFmtId')!);
      xf.setAttribute('applyNumberFormat', '1');
      changed = true;
    }
    if (changed) existing.setAttribute('s', String(appendStyle('cellXfs', xf)));
  }
  for (let i = 0; i < current.sheets.length; i++) {
    const sheet = current.sheets[i],
      old = baseline.sheets[i];
    const doc = parseXML(await zip.file(old.sourcePath!)!.async('string'));
    const root = doc.documentElement,
      data = ensure(root, 'sheetData');
    const nodes = new Map(elements(data, 'c').map((c) => [c.getAttribute('r')!, c]));
    let dirty = await writeWorksheetLinks(zip, old.sourcePath!, root, old.cells, sheet.cells);
    if (await writeSheetTables(zip, old.sourcePath!, root, old, sheet)) dirty = true;
    if (sheet.validationRanges && !equal(sheet.validationRanges, old.validationRanges))
      throw Error(
        'Edit validation through cell rules; changing source range metadata is not supported.',
      );
    if (writeValidationChanges(root, old.cells, sheet.cells)) dirty = true;
    if (
      (sheet.frozenRows || 0) !== (old.frozenRows || 0) ||
      (sheet.frozenColumns || 0) !== (old.frozenColumns || 0)
    ) {
      writeFrozenPanes(root, sheet.frozenRows || 0, sheet.frozenColumns || 0);
      dirty = true;
    }
    if (!equal(sheet.autoFilters, old.autoFilters)) {
      const previous = old.autoFilters || [],
        next = sheet.autoFilters || [];
      const changedTableFilter = (path: string | undefined, ref: string) =>
        sheet.tables?.some((t) => t.sourcePath === path && t.ref === ref && t.headerRows === 1);
      if (
        previous.some(
          (f) =>
            !next.some(
              (n) =>
                n.sourcePath === f.sourcePath &&
                (n.ref === f.ref || changedTableFilter(n.sourcePath, n.ref)),
            ),
        )
      )
        throw new Error(
          'Removing or resizing imported filter ranges is not supported. Clear their criteria instead.',
        );
      if (new Set(next.map((f) => f.sourcePath)).size !== next.length)
        throw new Error('Only one filter range per worksheet or table is supported.');
      for (const filter of next) {
        const before = previous.find((f) => f.sourcePath === filter.sourcePath);
        if (equal(before, filter)) continue;
        if (
          (before &&
            before.ref !== filter.ref &&
            !changedTableFilter(filter.sourcePath, filter.ref)) ||
          (!before &&
            filter.sourcePath !== sheet.sourcePath &&
            !changedTableFilter(filter.sourcePath, filter.ref))
        )
          throw new Error('Cannot change this imported filter range.');
        const filterDoc =
          filter.sourcePath === sheet.sourcePath
            ? doc
            : parseXML(await zip.file(filter.sourcePath!)!.async('string'));
        const filterRoot = filterDoc.documentElement;
        let node = child(filterRoot, 'autoFilter');
        if (!node) {
          node = filterDoc.createElementNS(filterRoot.namespaceURI, 'autoFilter');
          const beforeTags = [
            'sortState',
            'dataConsolidate',
            'customSheetViews',
            'mergeCells',
            'phoneticPr',
            'conditionalFormatting',
            'dataValidations',
            'hyperlinks',
            'printOptions',
            'pageMargins',
            'pageSetup',
            'headerFooter',
            'rowBreaks',
            'colBreaks',
            'drawing',
            'legacyDrawing',
            'tableParts',
            'extLst',
          ];
          filterRoot.insertBefore(
            node,
            Array.from(filterRoot.children).find((e) => beforeTags.includes(e.localName)) || null,
          );
        }
        node.setAttribute('ref', filter.ref);
        writeFilterColumns(node, filter.columns);
        if (filterDoc === doc) dirty = true;
        else zip.file(filter.sourcePath!, serialize(filterDoc), { createFolders: false });
      }
    }
    if (!!sheet.filterMode !== !!old.filterMode) {
      const properties = ensure(root, 'sheetPr');
      root.insertBefore(properties, root.firstChild);
      if (sheet.filterMode) properties.setAttribute('filterMode', '1');
      else properties.removeAttribute('filterMode');
      dirty = true;
    }
    const refs = new Set([...Object.keys(sheet.cells), ...Object.keys(old.cells)]);
    for (const ref of refs) {
      const before = old.cells[ref],
        after = sheet.cells[ref];
      const priorArray = old.arrayFormulas?.find((a) => a.anchor === ref);
      const nextArray = sheet.arrayFormulas?.find((a) => a.anchor === ref);
      const arrayChanged = !equal(priorArray, nextArray);
      if (equal(before, after) && !arrayChanged) continue;
      let node = nodes.get(ref);
      if (!node) {
        const [r, c] = coordinates(ref);
        let row = Array.from(data.children).find((e) => Number(e.getAttribute('r')) === r + 1);
        if (!row) {
          row = doc.createElementNS(root.namespaceURI || NS, 'row');
          row.setAttribute('r', String(r + 1));
          data.insertBefore(
            row,
            Array.from(data.children).find((e) => Number(e.getAttribute('r')) > r + 1) || null,
          );
        }
        node = doc.createElementNS(root.namespaceURI || NS, 'c');
        node.setAttribute('r', ref);
        row.insertBefore(
          node,
          Array.from(row.children).find((e) => coordinates(e.getAttribute('r')!)[1] > c) || null,
        );
        nodes.set(ref, node);
      }
      if (
        !after ||
        !before ||
        after.value !== before.value ||
        after.dataType !== before.dataType ||
        arrayChanged
      ) {
        const f = child(node, 'f');
        if (f || (after?.value.startsWith('=') && after.dataType !== 'text'))
          formulasChanged = true;
        if (f?.getAttribute('t') === 'array' && (!priorArray || !singleArray(priorArray)))
          throw new Error(`Editing array formula ${ref} requires an array-aware operation.`);
        // Expand a touched shared formula group before replacing one member.
        if (f?.getAttribute('t') === 'shared') {
          const shared = f.getAttribute('si');
          for (const [otherRef, other] of nodes) {
            const otherF = child(other, 'f');
            if (otherF?.getAttribute('t') === 'shared' && otherF.getAttribute('si') === shared) {
              otherF.removeAttribute('t');
              otherF.removeAttribute('si');
              otherF.removeAttribute('ref');
              otherF.textContent = old.cells[otherRef]?.value.slice(1) || '';
            }
          }
        }
        ['f', 'v', 'is'].forEach((tag) => remove(node!, tag));
        node.removeAttribute('t');
        const value = after?.value || '';
        if (value.startsWith('=') && after?.dataType !== 'text') {
          ensure(node, 'f').textContent = value.slice(1);
          if (nextArray) {
            const formula = ensure(node, 'f');
            formula.setAttribute('t', 'array');
            formula.setAttribute('ref', nextArray.ref);
          }
        } else if (
          value !== '' &&
          after?.dataType !== 'text' &&
          !value.startsWith("'") &&
          Number.isFinite(Number(value))
        )
          ensure(node, 'v').textContent = value;
        else if (after?.dataType === 'boolean') {
          node.setAttribute('t', 'b');
          ensure(node, 'v').textContent = value.toUpperCase() === 'TRUE' ? '1' : '0';
        } else if (after?.dataType === 'error') {
          node.setAttribute('t', 'e');
          ensure(node, 'v').textContent = value;
        } else if (value !== '') {
          node.setAttribute('t', 'inlineStr');
          const t = ensure(ensure(node, 'is'), 't');
          t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
          t.textContent = value.startsWith("'") ? value.slice(1) : value;
        }
        valuesChanged = true;
        changedCells.add(formulaCellIdentity(sheet.id, ref));
      }
      if (after) styleFor(node, before, after);
      dirty = true;
    }
    for (const [key, attr] of [
      ['rowHeights', 'ht'],
      ['hiddenRows', 'hidden'],
    ] as const)
      if (!equal(sheet[key], old[key])) {
        if (key === 'rowHeights')
          for (const index of Object.keys(sheet.rowHeights || {})) {
            const number = Number(index) + 1;
            if (!Array.from(data.children).some((e) => Number(e.getAttribute('r')) === number)) {
              const row = doc.createElementNS(root.namespaceURI || NS, 'row');
              row.setAttribute('r', String(number));
              data.insertBefore(
                row,
                Array.from(data.children).find((e) => Number(e.getAttribute('r')) > number) || null,
              );
            }
          }
        for (const row of Array.from(data.children)) {
          const index = Number(row.getAttribute('r')) - 1;
          if (key === 'hiddenRows') {
            if (sheet.hiddenRows?.includes(index)) row.setAttribute('hidden', '1');
            else row.removeAttribute('hidden');
          } else if (sheet.rowHeights?.[index] !== undefined) {
            row.setAttribute(attr, String(sheet.rowHeights[index] * 0.75));
            row.setAttribute('customHeight', '1');
          }
        }
        dirty = true;
      }
    if (!equal(sheet.columnWidths, old.columnWidths)) {
      const cols = ensure(root, 'cols');
      root.insertBefore(cols, data);
      for (const [index, width] of Object.entries(sheet.columnWidths || {}))
        if (width !== old.columnWidths?.[index]) {
          const n = Number(index) + 1;
          const existing = Array.from(cols.children).find(
            (e) => Number(e.getAttribute('min')) <= n && Number(e.getAttribute('max')) >= n,
          );
          const col =
            (existing?.cloneNode(true) as Element) ||
            doc.createElementNS(root.namespaceURI || NS, 'col');
          if (existing) {
            const min = Number(existing.getAttribute('min')),
              max = Number(existing.getAttribute('max'));
            if (min < n) {
              const left = existing.cloneNode(true) as Element;
              left.setAttribute('max', String(n - 1));
              cols.insertBefore(left, existing);
            }
            if (max > n) {
              const right = existing.cloneNode(true) as Element;
              right.setAttribute('min', String(n + 1));
              cols.insertBefore(right, existing);
            }
            existing.remove();
          }
          col.setAttribute('min', String(n));
          col.setAttribute('max', String(n));
          col.setAttribute('width', String((width - 5) / 7));
          col.setAttribute('customWidth', '1');
          cols.appendChild(col);
        }
      Array.from(cols.children)
        .sort((a, b) => Number(a.getAttribute('min')) - Number(b.getAttribute('min')))
        .forEach((e) => cols.appendChild(e));
      dirty = true;
    }
    if (sheet.state !== old.state) {
      const entry = mapping.sheets[i].element;
      entry.setAttribute('state', sheet.state || 'visible');
      workbookChanged = true;
    }
    if (dirty) zip.file(old.sourcePath!, serialize(doc), { createFolders: false });
  }
  if (
    valuesChanged ||
    current.sheets.some(
      (s, i) =>
        !equal(s.hiddenRows, baseline.sheets[i].hiddenRows) ||
        !!s.filterMode !== !!baseline.sheets[i].filterMode,
    )
  ) {
    changedCells = staticFormulaDependents(current.sheets, changedCells);
    // Recalculate supported expressions; leave unsupported cached results untouched and request Excel recalculation.
    for (const sheet of current.sheets) {
      const doc = parseXML(await zip.file(sheet.sourcePath!)!.async('string'));
      let dirty = false;
      for (const node of elements(doc, 'c'))
        if (child(node, 'f')) {
          const ref = node.getAttribute('r')!;
          if (!changedCells.has(formulaCellIdentity(sheet.id, ref))) continue;
          const result = calc.result(sheet, ref),
            value = result.value;
          if (result.kind === 'unsupported') continue;
          remove(node, 'v');
          node.removeAttribute('t');
          if (typeof value === 'string')
            node.setAttribute('t', result.kind === 'error' ? 'e' : 'str');
          if (typeof value === 'boolean') node.setAttribute('t', 'b');
          ensure(node, 'v').textContent =
            typeof value === 'boolean' ? (value ? '1' : '0') : String(value);
          dirty = true;
        }
      if (dirty) zip.file(sheet.sourcePath!, serialize(doc), { createFolders: false });
    }
    const calcPr = ensure(mapping.book.documentElement, 'calcPr');
    calcPr.setAttribute('fullCalcOnLoad', '1');
    calcPr.setAttribute('forceFullCalc', '1');
    workbookChanged = true;
    // Formula edits invalidate the previous calculation order. The chain is optional metadata.
    if (formulasChanged) {
      const relPath = 'xl/_rels/workbook.xml.rels';
      const rels = parseXML(await zip.file(relPath)!.async('string'));
      const types = parseXML(await zip.file('[Content_Types].xml')!.async('string'));
      let removed = false;
      for (const rel of elements(rels, 'Relationship')) {
        if (!rel.getAttribute('Type')?.endsWith('/calcChain')) continue;
        const path = resolvePart('xl/workbook.xml', rel.getAttribute('Target')!);
        zip.remove(path);
        rel.remove();
        for (const part of elements(types, 'Override'))
          if (part.getAttribute('PartName') === '/' + path) part.remove();
        removed = true;
      }
      if (removed) {
        zip.file(relPath, serialize(rels), { createFolders: false });
        zip.file('[Content_Types].xml', serialize(types), { createFolders: false });
      }
    }
  }
  if (stylesChanged) zip.file('xl/styles.xml', serialize(styles), { createFolders: false });
  if (workbookChanged)
    zip.file('xl/workbook.xml', serialize(mapping.book), { createFolders: false });
  return new Blob([await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}
function serializeElement(node: Element) {
  return new XMLSerializer().serializeToString(node);
}
