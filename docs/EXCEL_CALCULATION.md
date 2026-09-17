# Excel calculation compatibility

Ten financial functions now have 170 native Excel cases and a loan edit/undo/reload/cache-export browser workflow. See [EXCEL_FINANCIAL.md](EXCEL_FINANCIAL.md) for scope, numerical tolerance and remaining financial functions.

The engine now calculates absolute named references and formulas, contiguous table selectors, scalar indirect references and multiple-criteria functions. This is a tested subset of desktop Excel, not a complete calculation engine.

The scalar statistical extension adds 27 function names, with 114 native Excel cases and browser edit/reload/cache-export coverage. See [EXCEL_STATISTICS.md](EXCEL_STATISTICS.md) for exact scope, aliases and remaining limits.

The six core aggregates also have reference-aware coercion, blank/error semantics and bounded traversal, verified against 128 native Excel cases. See [EXCEL_AGGREGATES.md](EXCEL_AGGREGATES.md).

## Implemented behavior

- Workbook and sheet-local defined names, constants and formulas; sheet-qualified names; table names used as data ranges. Relative defined-name anchors and external workbooks remain unsupported.
- Imported table metadata, default data references, `#All`, `#Data`, `#Headers`, `#Totals`, `#This Row`/`@`, column selectors and contiguous column spans. Metadata persists in native backups. Table creation, resizing, calculated-column propagation, selector unions and full implicit-intersection semantics remain open.
- `ROW`, `COLUMN`, `ROWS`, `COLUMNS`; whole-column/whole-row reference geometry; bounded lazy `VLOOKUP`/`HLOOKUP` scans that avoid iterating over trailing empty cells.
- `INDIRECT` with A1 text, same-workbook named references and English R1C1 absolute/relative text. Invalid references and non-reference named constants produce `#REF!`; external sources are not fetched.
- `ISNA`, `ISERROR`, `ISERR`, `IFNA`, `NA`, error literals and omitted arguments. Unsupported work, exhausted limits and cycles cannot be hidden as successful error-handler results.
- `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, wildcard/escaped criteria, dimension validation, error criteria, numeric text boundaries and empty-cell criteria. `SUMIF`/`AVERAGEIF` use the value range's starting cell when its dimensions differ from the criteria range.
- Scalar `XLOOKUP` and `XMATCH`, including `_xlfn.` imports, exact/nearest/wildcard matching, forward/reverse search and sorted ascending/descending binary search. Binary wildcard search and multi-cell return spills remain explicitly unsupported. Sortedness is a precondition for binary search, as in Excel.
- Tested SUBTOTAL operations 1–11 and 101–111 with hidden/filter visibility, nested exclusions, reference arguments and error/count distinctions. [EXCEL_FILTERS.md](EXCEL_FILTERS.md) records 132 native subtotal results, six additional native filtering results and actual browser-export validation in Excel.

Names, table references, `INDIRECT` and visibility-sensitive SUBTOTAL make static dependency discovery incomplete. On export, these formulas are conservatively marked affected after value/visibility changes; ordinary downstream dependencies are then propagated. Supported caches are updated in the original package. Unsupported caches remain unchanged and Excel recalculation is requested. The engine memoizes parsed expressions, cell results, ordinary errors and lookup positions for one workbook snapshot.

## Independent evidence

`tests/fixtures/excel-formulas-input.json` contains authored data and 61 formulas. `scripts/verify-excel-formulas.ps1` creates a separate workbook in installed Microsoft Excel, calculates it, and records typed results in `tests/fixtures/excel-formulas-oracle.json`. The recorded application is Excel 16.0, build 4627, language 1031. The receipt contains the input SHA-256. Local generated XLSX evidence is `.local/excel-formulas-oracle.xlsx`.

The script uses native ranges for range-name definitions, disables macros and restores application settings. A bounded startup retry can dismiss the ordinary German activation reminder belonging to that test process; it does not change activation. German Excel needs localized error names and R1C1 syntax inside string arguments. The receipt records both the authored formula and the actual native formula, and converts numeric COM error codes to invariant Excel error tokens.

All 61 results are checked by `src/formula-references.test.ts`, alongside safety/scope regressions. `src/formula-export.test.ts` checks table metadata, native-model retention, byte-identical unchanged output, edited named/indirect/table caches, downstream recalculation and untouched package payloads. `tests/excel-calculation.spec.ts` exercises an actual input edit, persisted reload and XLSX download in Chromium.

`src/modern-lookup.test.ts` contains 19 scalar lookup/error cases. The installed desktop build predates XLOOKUP, so these are not presented as native XLOOKUP validation. The downloaded Apache POI `xlookup.xlsx` additionally supplies three matching scalar stored results; its two-column spill remains unsupported.

## Downloaded workbook audit

Run the opt-in audit after downloading the corpus:

```powershell
$env:NOFFICE_FORMULA_AUDIT = '1'
npm.cmd test -- src/formula-corpus-audit.test.ts
```

The report is written to `.local/formula-audit/report.json`. It compares every imported formula with its stored source cache using typed equality and numeric tolerance. Stored caches can be stale, so this is a separate evidence layer from fresh desktop recalculation. The audit asserts 7,274 matching formulas and zero unsupported/mismatched results in `StructuredRefs-lots-with-lookups.xlsx`, the three matching scalar modern lookups, and all 325 numeric/45 date format results. A successful audit run does not certify complete Excel behavior: every discrepancy and unsupported calculation remains in the report.

The subtotal/filter delivery's 20-workbook run completed in 164 seconds and passed all targeted assertions. Its totals are 8,060 formulas, 8,057 matching stored results, one unsupported calculation, zero mismatches and two absent caches. The earlier reference-only implementation had 7,736 matches, 313 unsupported results and nine mismatches. Importing the large structured workbook under jsdom dominates the audit; these timings are not a browser performance guarantee. All 20 XLSX import/export browser scenarios also passed; the existing strict whole-suite parity gate remains failing for the safely rejected deep DOCX and the protected workbook's missing edit scenario.

The remaining unsupported corpus formula is one dynamic XLOOKUP spill. All 18 subtotals now match. The nine text-format discrepancies and 294 unsupported numeric-format expressions are resolved for these fixtures; [NUMBER_FORMATS.md](NUMBER_FORMATS.md) describes the implementation, native evidence and remaining format limits. Full number-format/filter semantics, dynamic arrays, iterative calculation, relative named formulas, reference repair, function coverage and worker-based cancellation remain open.

Reference material: [Microsoft structured references](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables), [INDIRECT](https://support.microsoft.com/en-us/excel/functions/indirect-function), [XLOOKUP](https://support.microsoft.com/en-us/excel/functions/xlookup-function), and [SUMIF/SUMIFS range constraints](https://support.microsoft.com/en-us/excel/how-to-correct-a-value-error-in-the-sumif-sumifs-function).
