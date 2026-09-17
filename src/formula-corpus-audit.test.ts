import { it, expect } from 'vitest';
import { readFile, readdir, mkdir, writeFile } from 'node:fs/promises';
import { importWorkbook } from './xlsx-import';
import { calculator } from './formulas';

it.skipIf(!process.env.NOFFICE_FORMULA_AUDIT)(
  'audits formula results against external XLSX stored caches (not a fresh Excel oracle)',
  async () => {
    const root = '.local/office-corpus/input/excel';
    const report = [];
    for (const name of await readdir(root)) {
      if (!name.endsWith('.xlsx')) continue;
      if (process.env.NOFFICE_FORMULA_FILTER && !name.includes(process.env.NOFFICE_FORMULA_FILTER))
        continue;
      const input = await readFile(`${root}/${name}`);
      const started = performance.now();
      const content = await importWorkbook(Uint8Array.from(input).buffer);
      const imported = performance.now();
      const calc = calculator(content.sheets);
      let formulas = 0,
        equal = 0,
        unsupported = 0,
        noCache = 0;
      const mismatches = [],
        unsupportedExamples = [];
      for (const sheet of content.sheets)
        for (const [ref, cell] of Object.entries(sheet.cells)) {
          if (!cell.value.startsWith('=') || cell.dataType === 'text') continue;
          formulas++;
          const result = calc.result(sheet, ref);
          if (result.kind === 'unsupported') {
            unsupported++;
            if (unsupportedExamples.length < 5)
              unsupportedExamples.push({ ref, formula: cell.value, result });
            continue;
          }
          if (cell.cachedValue === undefined) {
            noCache++;
            continue;
          }
          const match =
            typeof cell.cachedValue === 'number' && typeof result.value === 'number'
              ? Math.abs(cell.cachedValue - result.value) <=
                1e-9 * Math.max(1, Math.abs(cell.cachedValue))
              : result.value === cell.cachedValue;
          if (match) equal++;
          else
            mismatches.push({
              sheet: sheet.name,
              ref,
              formula: cell.value,
              cached: cell.cachedValue,
              result,
            });
        }
      report.push({ name, formulas, equal, unsupported, noCache, mismatches, unsupportedExamples });
      console.log(
        `${name}: import ${Math.round(imported - started)}ms, calculate ${Math.round(performance.now() - imported)}ms, ${equal}/${formulas} cached results match`,
      );
    }
    await mkdir('.local/formula-audit', { recursive: true });
    await writeFile(
      `.local/formula-audit/${process.env.NOFFICE_FORMULA_FILTER || 'report'}.json`,
      JSON.stringify(report, null, 2),
    );
    console.log(
      report.map(({ mismatches, ...rest }) => ({ ...rest, mismatches: mismatches.length })),
    );
    expect(report).toHaveLength(process.env.NOFFICE_FORMULA_FILTER ? 1 : 20);
    const structured = report.find((r) => r.name === 'StructuredRefs-lots-with-lookups.xlsx');
    if (structured)
      expect(structured).toMatchObject({
        formulas: 7274,
        equal: 7274,
        unsupported: 0,
        mismatches: [],
      });
    const modern = report.find((r) => r.name === 'xlookup.xlsx');
    if (modern)
      expect(modern).toMatchObject({ formulas: 4, equal: 3, unsupported: 1, mismatches: [] });
    for (const [name, formulas] of [
      ['ConditionalFormattingSamples.xlsx', 352],
      ['NumberFormatTests.xlsx', 325],
      ['DateFormatTests.xlsx', 45],
    ] as const) {
      const formats = report.find((r) => r.name === name);
      if (formats)
        expect(formats).toMatchObject({
          formulas,
          equal: formulas,
          unsupported: 0,
          mismatches: [],
        });
    }
  },
  300000,
);
