import JSZip from 'jszip';
import { address, colName, coordinates, translateFormula } from './formulas';
import { elements, parseXML, workbookSheets } from './xlsx-import';
import {
  checkedStructureEdit,
  moveIndex,
  structuralFormula,
  structuralRange,
  type StructureEdit,
} from './sheet-structure';

const serialize = (doc: Document) => new XMLSerializer().serializeToString(doc);
const children = (node: Element, name: string) =>
  Array.from(node.children).filter((c) => c.localName === name);

function rangeList(value: string, edit: StructureEdit, inherit = false) {
  return value
    .trim()
    .split(/\s+/)
    .map((ref) => {
      if (inherit && edit.action === 'insert' && edit.at > 0) {
        const [a, b = a] = ref.split(':'),
          start = coordinates(a),
          end = coordinates(b);
        const axis = edit.axis === 'row' ? 0 : 1;
        // Excel copies validation from the row above / column to the left even when
        // inserting immediately after the previous rule's last row or column.
        if (end[axis] === edit.at - 1) {
          end[axis] += edit.count;
          return address(...start) + ':' + address(...end);
        }
      }
      return structuralRange(ref, edit);
    })
    .filter((ref) => ref !== '#REF!')
    .join(' ');
}

/** Operates on a copy. All unhandled reference-bearing features reject before any output is returned. */
export async function restructureXlsx(
  data: ArrayBuffer,
  edits: StructureEdit[],
): Promise<ArrayBuffer> {
  if (!edits.length) return data;
  if (edits.length > 100)
    throw Error(
      'This workbook has reached the 100 structural operations limit. Export and reopen it to continue.',
    );
  const zip = await JSZip.loadAsync(data),
    mapping = await workbookSheets(zip);
  if (
    Object.keys(zip.files).some((path) =>
      /^xl\/(?:tables|charts|pivotTables|pivotCache|externalLinks|threadedComments)\//.test(path),
    )
  )
    throw Error(
      'Row and column operations in workbooks with tables, charts, pivots, external links or threaded comments are not supported yet.',
    );
  if (elements(mapping.book, 'workbookProtection').length)
    throw Error('This workbook is protected.');
  const documents = new Map<string, Document>();
  const originalXml = new Map<string, string>();
  for (const s of mapping.sheets) {
    const entry = zip.file(s.path);
    if (!entry || !/\/worksheets\//.test(s.path))
      throw Error('Only standard worksheets support row and column operations.');
    const doc = parseXML(await entry.async('string'));
    documents.set(s.name, doc);
    originalXml.set(s.name, serialize(doc));
  }
  for (const edit of edits) {
    checkedStructureEdit(edit);
    const target = documents.get(edit.sheet);
    if (!target) throw Error('The worksheet no longer exists.');
    // Drawings/notes have separate anchors, and extension payloads can hold additional references.
    const blocked = [
      'sheetProtection',
      'drawing',
      'legacyDrawing',
      'legacyDrawingHF',
      'oleObjects',
      'controls',
      'customSheetViews',
      'scenarios',
      'dataConsolidate',
      'cellWatches',
      'smartTags',
      'protectedRanges',
    ];
    if (blocked.some((tag) => elements(target, tag).length))
      throw Error(
        'This worksheet contains protection, drawings, notes or other anchors that cannot yet be repaired by row and column operations.',
      );
    for (const doc of documents.values())
      if (elements(doc, 'extLst').length)
        throw Error(
          'Worksheet extension references cannot yet be repaired by row and column operations.',
        );
    for (const s of mapping.sheets) {
      const doc = documents.get(s.name)!;
      // Expand shared groups while their original anchors and offsets are still available.
      const shared = new Map<string, { ref: string; formula: string }>();
      for (const f of elements(doc, 'f')) {
        if (f.getAttribute('t') && !['normal', 'shared'].includes(f.getAttribute('t')!))
          throw Error(
            'Array and data-table formulas do not yet support row and column operations.',
          );
        if (f.getAttribute('t') === 'shared' && f.textContent)
          shared.set(f.getAttribute('si')!, {
            ref: f.parentElement!.getAttribute('r')!,
            formula: f.textContent,
          });
      }
      for (const f of elements(doc, 'f')) {
        if (f.getAttribute('t') === 'shared') {
          const master = shared.get(f.getAttribute('si')!);
          if (!master) throw Error('Invalid shared formula group.');
          const [r, c] = coordinates(f.parentElement!.getAttribute('r')!),
            [mr, mc] = coordinates(master.ref);
          f.textContent = translateFormula('=' + master.formula, r - mr, c - mc).slice(1);
          for (const a of ['t', 'si', 'ref']) f.removeAttribute(a);
        }
        f.textContent = structuralFormula(f.textContent || '', s.name, edit);
        children(f.parentElement!, 'v').forEach((v) => v.remove());
        f.parentElement!.removeAttribute('t');
      }
      for (const tag of ['formula', 'formula1', 'formula2'])
        for (const f of elements(doc, tag))
          f.textContent = structuralFormula(f.textContent || '', s.name, edit);
      // Excel keeps hyperlink destinations as literal strings during structural edits.
      // Their cell anchors move below, but their location/relationship targets stay unchanged.
    }
    for (const name of elements(mapping.book, 'definedName')) {
      const local = name.getAttribute('localSheetId');
      const context = local === null ? '' : mapping.sheets[Number(local)]?.name || '';
      const formula = name.textContent || '';
      if (!context && /(?:^|[,(])\s*\$?[A-Z]{1,3}\$?[1-9]\d*/i.test(formula))
        throw Error('Workbook names with unqualified relative references cannot yet be repaired.');
      name.textContent = structuralFormula(formula, context, edit);
    }
    const sheetData = elements(target, 'sheetData')[0];
    if (!sheetData) throw Error('Worksheet cell data is missing.');
    const rows = children(sheetData, 'row');
    const existingCellCount = elements(sheetData, 'c').length;
    const extraCellCount =
      edit.action === 'insert'
        ? edit.count *
          (edit.axis === 'row'
            ? children(
                rows.find((row) => Number(row.getAttribute('r')) === edit.at) || sheetData,
                'c',
              ).length
            : rows.length)
        : 0;
    if (existingCellCount + extraCellCount > 100000)
      throw Error(
        'Structural operations are currently limited to 100,000 represented cells per worksheet.',
      );
    // Copy only formatting into inserted cells. Contents, formulas and hyperlinks are never duplicated.
    const formatCell = (source: Element, ref: string) => {
      const c = target.createElementNS(source.namespaceURI, 'c');
      c.setAttribute('r', ref);
      if (source.hasAttribute('s')) c.setAttribute('s', source.getAttribute('s')!);
      return c;
    };
    const insertedRows: Element[] = [];
    if (edit.axis === 'row' && edit.action === 'insert') {
      const previous = rows.find((row) => Number(row.getAttribute('r')) === edit.at);
      if (previous)
        for (let i = 0; i < edit.count; i++) {
          const row = previous.cloneNode(false) as Element;
          row.setAttribute('r', String(edit.at + i + 1));
          row.removeAttribute('spans');
          // Hidden/outline state belongs to the existing rows.
          for (const a of ['hidden', 'collapsed', 'outlineLevel']) row.removeAttribute(a);
          for (const c of children(previous, 'c'))
            if (c.hasAttribute('s'))
              row.appendChild(
                formatCell(c, address(edit.at + i, coordinates(c.getAttribute('r')!)[1])),
              );
          insertedRows.push(row);
        }
    }
    for (const row of rows) {
      const oldRow = Number(row.getAttribute('r')) - 1;
      const newRow = edit.axis === 'row' ? moveIndex(oldRow, edit) : oldRow;
      if (newRow === null) {
        row.remove();
        continue;
      }
      if (newRow >= 10000)
        throw Error('This operation would exceed the editable limit of 10,000 rows.');
      row.setAttribute('r', String(newRow + 1));
      row.removeAttribute('spans');
      const cells = children(row, 'c');
      const previous =
        edit.axis === 'column' && edit.action === 'insert'
          ? cells.find((c) => coordinates(c.getAttribute('r')!)[1] === edit.at - 1)
          : undefined;
      for (const cell of cells) {
        const oldCol = coordinates(cell.getAttribute('r')!)[1];
        const newCol = edit.axis === 'column' ? moveIndex(oldCol, edit) : oldCol;
        if (newCol === null) {
          cell.remove();
          continue;
        }
        if (newCol >= 256)
          throw Error('This operation would exceed the editable limit of 256 columns.');
        cell.setAttribute('r', address(newRow, newCol));
      }
      if (previous?.hasAttribute('s'))
        for (let i = 0; i < edit.count; i++)
          row.appendChild(formatCell(previous, address(newRow, edit.at + i)));
      children(row, 'c')
        .sort((a, b) => coordinates(a.getAttribute('r')!)[1] - coordinates(b.getAttribute('r')!)[1])
        .forEach((c) => row.appendChild(c));
    }
    insertedRows.forEach((row) => sheetData.appendChild(row));
    children(sheetData, 'row')
      .sort((a, b) => Number(a.getAttribute('r')) - Number(b.getAttribute('r')))
      .forEach((row) => sheetData.appendChild(row));
    if (edit.axis === 'column') {
      const cols = elements(target, 'col');
      const previous = cols.find(
        (c) => Number(c.getAttribute('min')) <= edit.at && Number(c.getAttribute('max')) >= edit.at,
      );
      const copy = previous?.cloneNode(false) as Element | undefined;
      for (const col of cols) {
        const ref = `${colName(Number(col.getAttribute('min')) - 1)}:${colName(Number(col.getAttribute('max')) - 1)}`;
        const moved = structuralRange(ref, edit);
        if (moved === '#REF!') col.remove();
        else {
          const [a, b] = moved.split(':');
          col.setAttribute('min', String(coordinates(a + '1')[1] + 1));
          col.setAttribute('max', String(coordinates(b + '1')[1] + 1));
        }
      }
      // Insertion inside an existing column span already expands that span.
      if (
        copy &&
        edit.action === 'insert' &&
        !elements(target, 'col').some(
          (c) =>
            Number(c.getAttribute('min')) <= edit.at + 1 &&
            Number(c.getAttribute('max')) >= edit.at + 1,
        )
      ) {
        copy.setAttribute('min', String(edit.at + 1));
        copy.setAttribute('max', String(edit.at + edit.count));
        for (const a of ['hidden', 'collapsed', 'outlineLevel']) copy.removeAttribute(a);
        elements(target, 'cols')[0].appendChild(copy);
      }
      const parent = elements(target, 'cols')[0];
      if (parent)
        children(parent, 'col')
          .sort((a, b) => Number(a.getAttribute('min')) - Number(b.getAttribute('min')))
          .forEach((c) => parent.appendChild(c));
    }
    for (const node of elements(target, 'autoFilter')) {
      const old = node.getAttribute('ref')!;
      if (edit.axis === 'column') {
        const start = coordinates(old.split(':')[0])[1];
        const newRange = structuralRange(old, edit);
        if (newRange !== '#REF!')
          for (const filter of children(node, 'filterColumn')) {
            const col = moveIndex(start + Number(filter.getAttribute('colId')), edit);
            if (col === null) filter.remove();
            else filter.setAttribute('colId', String(col - coordinates(newRange.split(':')[0])[1]));
          }
      }
    }
    for (const tag of [
      'dimension',
      'mergeCell',
      'autoFilter',
      'sortState',
      'sortCondition',
      'hyperlink',
    ]) {
      for (const node of elements(target, tag)) {
        if (!node.hasAttribute('ref')) continue;
        const ref = structuralRange(node.getAttribute('ref')!, edit);
        if (ref === '#REF!') {
          if (tag === 'dimension') node.setAttribute('ref', 'A1');
          else node.remove();
        } else node.setAttribute('ref', ref);
      }
    }
    for (const tag of ['dataValidation', 'conditionalFormatting', 'ignoredError', 'selection']) {
      for (const node of elements(target, tag)) {
        const refs = node.getAttribute('sqref');
        if (refs) {
          const moved = rangeList(refs, edit, tag === 'dataValidation');
          if (moved) node.setAttribute('sqref', moved);
          else if (tag === 'selection') node.setAttribute('sqref', 'A1');
          else {
            node.remove();
            continue;
          }
        }
        if (node.hasAttribute('activeCell')) {
          const moved = structuralRange(node.getAttribute('activeCell')!, edit);
          node.setAttribute('activeCell', moved === '#REF!' ? 'A1' : moved);
          node.removeAttribute('activeCellId');
        }
      }
    }
    for (const view of [...elements(target, 'sheetView'), ...elements(target, 'pane')]) {
      if (view.hasAttribute('topLeftCell')) {
        const moved = structuralRange(view.getAttribute('topLeftCell')!, edit);
        view.setAttribute('topLeftCell', moved === '#REF!' ? 'A1' : moved);
      }
      if (view.localName === 'pane') {
        if (view.getAttribute('state') !== 'frozen')
          throw Error('Split panes cannot yet be repaired by row and column operations.');
        const attr = edit.axis === 'row' ? 'ySplit' : 'xSplit';
        const size = Number(view.getAttribute(attr) || 0);
        if (size && edit.at < size)
          view.setAttribute(
            attr,
            String(
              edit.action === 'insert'
                ? size + edit.count
                : size - Math.min(edit.count, size - edit.at),
            ),
          );
        if (!Number(view.getAttribute('xSplit')) && !Number(view.getAttribute('ySplit')))
          view.remove();
      }
    }
    for (const tag of ['rowBreaks', 'colBreaks']) {
      if ((edit.axis === 'row') !== (tag === 'rowBreaks')) continue;
      for (const parent of elements(target, tag))
        for (const brk of children(parent, 'brk')) {
          const moved = moveIndex(Number(brk.getAttribute('id')) - 1, edit);
          if (moved === null) brk.remove();
          else brk.setAttribute('id', String(moved + 1));
        }
    }
    for (const tag of ['mergeCells', 'dataValidations', 'rowBreaks', 'colBreaks'])
      for (const parent of elements(target, tag)) {
        if (!parent.children.length) parent.remove();
        else {
          parent.setAttribute('count', String(parent.children.length));
          if (parent.hasAttribute('manualBreakCount'))
            parent.setAttribute(
              'manualBreakCount',
              String(
                Array.from(parent.children).filter((c) => c.getAttribute('man') === '1').length,
              ),
            );
        }
      }
  }
  for (const s of mapping.sheets) {
    const xml = serialize(documents.get(s.name)!);
    if (xml !== originalXml.get(s.name)) zip.file(s.path, xml, { createFolders: false });
  }
  let calc = elements(mapping.book, 'calcPr')[0];
  if (!calc) {
    calc = mapping.book.createElementNS(mapping.book.documentElement.namespaceURI, 'calcPr');
    const next = Array.from(mapping.book.documentElement.children).find((e) =>
      ['oleSize', 'customWorkbookViews', 'extLst'].includes(e.localName),
    );
    mapping.book.documentElement.insertBefore(calc, next || null);
  }
  calc.setAttribute('fullCalcOnLoad', '1');
  calc.setAttribute('forceFullCalc', '1');
  zip.file('xl/workbook.xml', serialize(mapping.book), { createFolders: false });
  // A structural edit invalidates the old calculation order and cell caches throughout the workbook.
  zip.remove('xl/calcChain.xml');
  for (const path of ['xl/_rels/workbook.xml.rels', '[Content_Types].xml']) {
    const doc = parseXML(await zip.file(path)!.async('string'));
    let changed = false;
    for (const node of Array.from(doc.documentElement.children))
      if (
        node.getAttribute('Type')?.endsWith('/calcChain') ||
        node.getAttribute('PartName') === '/xl/calcChain.xml'
      ) {
        node.remove();
        changed = true;
      }
    if (changed) zip.file(path, serialize(doc), { createFolders: false });
  }
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
