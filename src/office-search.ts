import type { WorkbookContent, DeckContent } from './model';
import { calculator, coordinates, displayValue } from './formulas';
import { validationError } from './cell-validation';
import { textMatches, replaceMatches, type SearchOptions, type TextMatch } from './text-search';

export type CellMatch = { sheetId: string; ref: string; matches: TextMatch[] };
export function nextWorkbookMatch(
  content: WorkbookContent,
  matches: CellMatch[],
  sheetId: string,
  ref: string,
  previous = false,
  byColumns = false,
) {
  const sheetOrder = new Map(content.sheets.map((s, i) => [s.id, i]));
  const [row, column] = coordinates(ref);
  const compare = (match: CellMatch) => {
    const difference = sheetOrder.get(match.sheetId)! - sheetOrder.get(sheetId)!;
    if (difference) return difference;
    const [r, c] = coordinates(match.ref);
    return byColumns ? c - column || r - row : r - row || c - column;
  };
  return previous
    ? matches.filter((m) => compare(m) < 0).at(-1) || matches.at(-1)
    : matches.find((m) => compare(m) > 0) || matches[0];
}
export type WorkbookSearch = SearchOptions & {
  sheetId?: string;
  refs?: string[];
  lookIn?: 'formulas' | 'values' | 'notes';
  byColumns?: boolean;
};
export function workbookMatches(
  content: WorkbookContent,
  query: string,
  options: WorkbookSearch,
): CellMatch[] {
  const found: CellMatch[] = [],
    budget = { characters: 0, matches: 0 },
    calc = calculator(content.sheets);
  const selection = options.refs ? new Set(options.refs) : null;
  for (const sheet of content.sheets) {
    if (options.sheetId && sheet.id !== options.sheetId) continue;
    // Hidden sheets/rows cannot be navigated in this editor; do not silently edit them.
    if (sheet.state && sheet.state !== 'visible') continue;
    const rows = new Set(sheet.hiddenRows),
      columns = new Set(sheet.hiddenColumns);
    const entries = Object.entries(sheet.cells).sort(([a], [b]) => {
      const [ar, ac] = coordinates(a),
        [br, bc] = coordinates(b);
      return options.byColumns ? ac - bc || ar - br : ar - br || ac - bc;
    });
    for (const [ref, cell] of entries) {
      if (selection && !selection.has(ref)) continue;
      const [r, c] = coordinates(ref);
      if (rows.has(r) || columns.has(c)) continue;
      const text =
        options.lookIn === 'notes'
          ? cell.note || ''
          : options.lookIn === 'values'
            ? displayValue(calc(sheet, ref), cell)
            : cell.value;
      const matches = textMatches(text, query, options, budget);
      if (matches.length) found.push({ sheetId: sheet.id, ref, matches });
    }
  }
  return found;
}

/** Validate every edit first. A protected or invalid target cancels the whole operation. */
export function replaceWorkbookMatches(
  content: WorkbookContent,
  found: CellMatch[],
  replacement: string,
) {
  const edits = new Map<string, Map<string, string>>();
  for (const match of found) {
    const sheet = content.sheets.find((s) => s.id === match.sheetId)!,
      cell = sheet.cells[match.ref];
    if (sheet.protected && cell.locked !== false)
      throw Error(`${sheet.name}!${match.ref} is protected. No cells were changed.`);
    const value = replaceMatches(cell.value, match.matches, replacement, 32767);
    const error = validationError(cell, value, sheet, content.sheets, match.ref);
    if (error) throw Error(`${sheet.name}!${match.ref}: ${error} No cells were changed.`);
    if (value === cell.value) continue;
    if (!edits.has(sheet.id)) edits.set(sheet.id, new Map());
    edits.get(sheet.id)!.set(match.ref, value);
  }
  return {
    ...content,
    sheets: content.sheets.map((sheet) => {
      const changes = edits.get(sheet.id);
      if (!changes) return sheet;
      const cells = { ...sheet.cells };
      for (const [ref, value] of changes)
        cells[ref] = { ...cells[ref], value, dataType: undefined, cachedValue: undefined };
      return { ...sheet, cells };
    }),
  };
}

export type SlideMatch = { slideId: string; elementId: string; from: number; to: number };
export function nextPresentationMatch(
  content: DeckContent,
  matches: SlideMatch[],
  slideId: string,
  elementId: string | null,
  last: SlideMatch | undefined,
  previous = false,
) {
  const slideIndex = content.slides.findIndex((s) => s.id === slideId);
  const elementIndex =
    content.slides[slideIndex]?.elements.findIndex((e) => e.id === elementId) ?? -1;
  const offset = last?.slideId === slideId && last.elementId === elementId ? last.from : -1;
  const compare = (match: SlideMatch) => {
    const index = content.slides.findIndex((s) => s.id === match.slideId);
    return (
      index - slideIndex ||
      content.slides[index].elements.findIndex((e) => e.id === match.elementId) - elementIndex ||
      match.from - offset
    );
  };
  return previous
    ? matches.filter((m) => compare(m) < 0).at(-1) || matches.at(-1)
    : matches.find((m) => compare(m) > 0) || matches[0];
}
export function presentationMatches(
  content: DeckContent,
  query: string,
  options: SearchOptions = {},
): SlideMatch[] {
  const found: SlideMatch[] = [],
    budget = { characters: 0, matches: 0 };
  for (const slide of content.slides)
    for (const element of slide.elements) {
      if (element.type === 'image') continue;
      found.push(
        ...textMatches(element.text, query, options, budget).map((m) => ({
          ...m,
          slideId: slide.id,
          elementId: element.id,
        })),
      );
    }
  return found;
}
export function replacePresentationMatches(
  content: DeckContent,
  found: SlideMatch[],
  replacement: string,
): DeckContent {
  return {
    ...content,
    slides: content.slides.map((slide) => ({
      ...slide,
      elements: slide.elements.map((element) => {
        const matches = found.filter((m) => m.slideId === slide.id && m.elementId === element.id);
        return matches.length
          ? { ...element, text: replaceMatches(element.text, matches, replacement, 100000) }
          : element;
      }),
    })),
  };
}
