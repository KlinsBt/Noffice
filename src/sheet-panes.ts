import type { Sheet } from './model';
import { address, coordinates } from './formulas';

export type FreezeCommand = 'selection' | 'row' | 'column' | 'none';
export function freezeSheet(sheet: Sheet, command: FreezeCommand, active: string): Sheet {
  const [row, column] = coordinates(active);
  const frozenRows = command === 'selection' ? row : command === 'row' ? 1 : 0;
  const frozenColumns = command === 'selection' ? column : command === 'column' ? 1 : 0;
  if (command === 'selection' && !row && !column)
    throw Error('Select a cell below the rows or to the right of the columns you want to freeze.');
  if (frozenRows >= sheet.rows || frozenColumns >= sheet.cols)
    throw Error('Keep at least one row and column available to scroll.');
  for (const merge of sheet.merges || []) {
    const [start, end = start] = merge.split(':');
    const [r, c] = coordinates(start),
      [rr, cc] = coordinates(end);
    if ((r < frozenRows && rr >= frozenRows) || (c < frozenColumns && cc >= frozenColumns))
      throw Error(
        'The freeze boundary crosses a merged cell. Select a boundary outside the merged range.',
      );
  }
  return { ...sheet, frozenRows, frozenColumns };
}

/** Patch the represented worksheet view, retaining other windows and unrelated view metadata. */
export function writeFrozenPanes(root: Element, rows: number, columns: number) {
  if (
    !Number.isInteger(rows) ||
    !Number.isInteger(columns) ||
    rows < 0 ||
    rows > 10000 ||
    columns < 0 ||
    columns > 256
  )
    throw Error('Invalid frozen pane dimensions.');
  const direct = (el: Element, name: string) =>
    Array.from(el.children).find((c) => c.localName === name);
  const make = (name: string) => root.ownerDocument.createElementNS(root.namespaceURI, name);
  let views = direct(root, 'sheetViews');
  if (!views) {
    views = make('sheetViews');
    root.insertBefore(
      views,
      Array.from(root.children).find((c) => !['sheetPr', 'dimension'].includes(c.localName)) ||
        null,
    );
  }
  let view = direct(views, 'sheetView');
  if (!view) {
    view = make('sheetView');
    view.setAttribute('workbookViewId', '0');
    views.appendChild(view);
  }
  for (const child of Array.from(view.children))
    if (['pane', 'selection'].includes(child.localName)) child.remove();
  const selection = make('selection');
  const ref = rows || columns ? address(rows, columns) : 'A1';
  selection.setAttribute('activeCell', ref);
  selection.setAttribute('sqref', ref);
  view.insertBefore(selection, view.firstChild);
  if (rows || columns) {
    const pane = make('pane');
    if (rows) pane.setAttribute('ySplit', String(rows));
    if (columns) pane.setAttribute('xSplit', String(columns));
    const activePane = rows && columns ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
    pane.setAttribute('state', 'frozen');
    pane.setAttribute('topLeftCell', ref);
    pane.setAttribute('activePane', activePane);
    selection.setAttribute('pane', activePane);
    view.insertBefore(pane, selection);
  }
}
