import { describe, expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import numeric from '../tests/fixtures/filter-numeric-input.json';
import wildcard from '../tests/fixtures/filter-wildcard-input.json';
import numericOracle from '../tests/fixtures/filter-numeric-oracle.json';
import wildcardOracle from '../tests/fixtures/filter-wildcard-oracle.json';
import { applySheetFilters, readAutoFilters, writeFilterColumns } from './sheet-filters';
import { calculator } from './formulas';
import { parseXML } from './xlsx-import';
import type { Sheet } from './model';

for (const [name, spec, oracle] of [
  ['numeric', numeric, numericOracle],
  ['wildcard', wildcard, wildcardOracle],
] as const)
  describe(`native Excel ${name} filter`, () => {
    const fixture = (): Sheet => ({
      id: 'filter',
      name: spec.sheet,
      rows: 100,
      cols: 26,
      cells: Object.fromEntries([
        ...Object.entries(spec.cells).map(([ref, value]) => [
          ref,
          {
            value: String(value),
            ...(typeof value === 'string' && !value.startsWith('=') ? { dataType: 'text' } : {}),
          },
        ]),
        ...spec.cases.map(({ ref, formula }) => [ref, { value: formula }]),
      ]),
    });
    const run = () => {
      const sheet = fixture();
      return applySheetFilters(
        sheet,
        [sheet],
        [
          {
            ref: 'A1:B9',
            columns: [
              {
                col: name === 'numeric' ? 0 : 1,
                custom: [
                  {
                    operator: name === 'numeric' ? 'greaterThanOrEqual' : 'equal',
                    value: name === 'numeric' ? '4' : 'keep*',
                  },
                ],
              },
            ],
          },
        ],
      );
    };
    it('records exact input and reproduces hidden rows', () => {
      expect(
        createHash('sha256')
          .update(readFileSync(`tests/fixtures/filter-${name}-input.json`))
          .digest('hex'),
      ).toBe(oracle.inputSha256);
      expect(run().hiddenRows).toEqual(oracle.hiddenRows);
    });
    it.each(oracle.cases)('$formula', ({ ref, value, kind }) => {
      const sheet = run();
      expect(calculator([sheet]).result(sheet, ref)).toEqual({ value, kind });
    });
  });

it('combines columns, clears criteria and rejects unsupported filter types without mutation', () => {
  const sheet: Sheet = {
    id: 'f',
    name: 'Data',
    rows: 100,
    cols: 26,
    cells: {
      A1: { value: 'Amount' },
      B1: { value: 'Group' },
      A2: { value: '2' },
      B2: { value: 'keep' },
      A3: { value: '6' },
      B3: { value: 'keep' },
      A4: { value: '8' },
      B4: { value: 'drop' },
    },
  };
  const next = applySheetFilters(
    sheet,
    [sheet],
    [
      {
        ref: 'A1:B4',
        columns: [
          { col: 0, custom: [{ operator: 'greaterThan', value: '4' }] },
          { col: 1, values: ['keep'] },
        ],
      },
    ],
  );
  expect(next.hiddenRows).toEqual([1, 3]);
  expect(sheet.hiddenRows).toBeUndefined();
  const cleared = applySheetFilters(
    next,
    [next],
    next.autoFilters!.map((f) => ({ ...f, columns: [] })),
  );
  expect(cleared.hiddenRows).toEqual([]);
  expect(cleared.filterMode).toBe(false);
  expect(() =>
    applySheetFilters(sheet, [sheet], [{ ref: 'A1:B4', columns: [{ col: 1, unsupported: true }] }]),
  ).toThrow('cannot be reapplied');
  expect(() => applySheetFilters({ ...sheet, protected: true }, [sheet], [])).toThrow('protected');
});
it('retains unsupported predicates and roundtrips supported columns without removing sort metadata', () => {
  const doc = parseXML(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><autoFilter ref="A1:B9"><filterColumn colId="1"><iconFilter iconSet="3TrafficLights1" iconId="1"/></filterColumn><sortState ref="A1:B9"/></autoFilter></worksheet>',
  );
  expect(readAutoFilters(doc.documentElement)[0].columns[0].unsupported).toBe(true);
  const columns = [{ col: 1, values: ['<safe & quoted>'], blank: true }];
  writeFilterColumns(doc.documentElement.firstElementChild!, columns);
  expect(readAutoFilters(doc.documentElement)[0].columns).toEqual(columns);
  expect(doc.getElementsByTagName('sortState')).toHaveLength(1);
});
