import { expect, it } from 'vitest';
import {
  workbookMatches,
  nextWorkbookMatch,
  replaceWorkbookMatches,
  presentationMatches,
  nextPresentationMatch,
  replacePresentationMatches,
} from './office-search';
import { textElement, type WorkbookContent, type DeckContent } from './model';
const workbook = (): WorkbookContent => ({
  kind: 'excel',
  sheets: [
    {
      id: 'a',
      name: 'First',
      rows: 10,
      cols: 3,
      cells: {
        A1: { value: 'Atlas', bold: true },
        B1: { value: 'Atlas B' },
        A2: { value: 'Atlas A' },
        C1: { value: '=1+1', note: 'Atlas note' },
      },
    },
    { id: 'b', name: 'Second', rows: 10, cols: 3, cells: { A1: { value: 'atlas', italic: true } } },
  ],
});

it('searches sheets, formula results, notes and selection in row/column order', () => {
  const content = workbook();
  expect(workbookMatches(content, 'atlas', {}).map((m) => [m.sheetId, m.ref])).toEqual([
    ['a', 'A1'],
    ['a', 'B1'],
    ['a', 'A2'],
    ['b', 'A1'],
  ]);
  expect(
    workbookMatches(content, 'atlas', { sheetId: 'a', byColumns: true }).map((m) => m.ref),
  ).toEqual(['A1', 'A2', 'B1']);
  expect(workbookMatches(content, '2', { lookIn: 'values' }).map((m) => m.ref)).toEqual(['C1']);
  expect(workbookMatches(content, 'Atlas', { lookIn: 'notes' })).toHaveLength(1);
  expect(workbookMatches(content, 'Atlas', { wholeCell: true, matchCase: true })).toHaveLength(1);
  expect(
    workbookMatches(content, 'Atlas', { sheetId: 'a', refs: ['B1'] }).map((m) => m.ref),
  ).toEqual(['B1']);
});
it('navigates from a nonmatching cell in row/column order and wraps across sheets', () => {
  const content = workbook(),
    rows = workbookMatches(content, 'Atlas', {});
  expect(nextWorkbookMatch(content, rows, 'a', 'C1')?.ref).toBe('A2');
  expect(nextWorkbookMatch(content, rows, 'a', 'C1', true)?.ref).toBe('B1');
  expect(nextWorkbookMatch(content, rows, 'b', 'B2')?.sheetId).toBe('a');
  expect(nextWorkbookMatch(content, rows, 'a', 'A1', true)?.sheetId).toBe('b');
  const columns = workbookMatches(content, 'Atlas', { byColumns: true });
  expect(nextWorkbookMatch(content, columns, 'a', 'A3', false, true)?.ref).toBe('B1');
});

it('replaces atomically, retains cell formatting and invalidates changed formula caches', () => {
  const content = workbook();
  const next = replaceWorkbookMatches(content, workbookMatches(content, 'Atlas', {}), 'Orion');
  expect(next.sheets[0].cells.A1).toMatchObject({ value: 'Orion', bold: true });
  expect(next.sheets[1].cells.A1).toMatchObject({ value: 'Orion', italic: true });
  expect(content.sheets[0].cells.A1.value).toBe('Atlas');
  content.sheets[1].protected = true;
  expect(() =>
    replaceWorkbookMatches(content, workbookMatches(content, 'Atlas', {}), 'Orion'),
  ).toThrow('protected');
  expect(content.sheets[0].cells.A1.value).toBe('Atlas');
  content.sheets[0].cells.C1.cachedValue = 2;
  expect(
    replaceWorkbookMatches(content, workbookMatches(content, '1', { sheetId: 'a' }), '3').sheets[0]
      .cells.C1,
  ).toMatchObject({ value: '=3+3', cachedValue: undefined });
});
it('rejects validation and size violations and skips invisible cells explicitly', () => {
  const content = workbook();
  content.sheets[0].cells.A1.validation = {
    type: 'list',
    formulae: ['"Atlas,Other"'],
    showErrorMessage: true,
    errorStyle: 'stop',
  };
  expect(() =>
    replaceWorkbookMatches(content, workbookMatches(content, 'Atlas', {}), 'Orion'),
  ).toThrow('No cells were changed');
  delete content.sheets[0].cells.A1.validation;
  expect(() =>
    replaceWorkbookMatches(content, workbookMatches(content, 'Atlas', {}), 'x'.repeat(32768)),
  ).toThrow('limit');
  content.sheets[0].hiddenRows = [0];
  content.sheets[1].state = 'hidden';
  expect(workbookMatches(content, 'Atlas', {}).map((m) => m.ref)).toEqual(['A2']);
});
it('replaces slide occurrences without losing object identities, geometry or notes', () => {
  const element = textElement('Atlas Atlas', { bold: true, x: 81 });
  const deck: DeckContent = {
    kind: 'powerpoint',
    slides: [{ id: 's', background: '#fff', notes: 'Atlas notes', elements: [element] }],
  };
  const found = presentationMatches(deck, 'Atlas', { wholeWord: true });
  expect(found).toHaveLength(2);
  expect(nextPresentationMatch(deck, found, 's', element.id, undefined)).toEqual(found[0]);
  expect(nextPresentationMatch(deck, found, 's', element.id, found[0])).toEqual(found[1]);
  expect(nextPresentationMatch(deck, found, 's', element.id, undefined, true)).toEqual(found[1]);
  const next = replacePresentationMatches(deck, [found[1]], 'Orion');
  expect(next.slides[0].elements[0]).toMatchObject({
    id: element.id,
    text: 'Atlas Orion',
    bold: true,
    x: 81,
  });
  expect(next.slides[0].notes).toBe('Atlas notes');
  expect(element.text).toBe('Atlas Atlas');
  expect(() => replacePresentationMatches(deck, found, 'x'.repeat(100000))).toThrow('limit');
});
