# Aggregate calculation compatibility

SUM, AVERAGE, MIN, MAX, COUNT and COUNTA now retain the difference between direct arguments, referenced cells, empty cells and formulas returning empty text. `aggregate.ts` owns numeric/coercion behavior; `formulas.ts` supplies reference-aware iteration and error values.

Examples corrected against native Excel:

| Formula            | Result    |
| ------------------ | --------- |
| `SUM(TRUE,"3",2)`  | 6         |
| `COUNTA(1,"a","")` | 3         |
| `AVERAGE("a")`     | `#VALUE!` |
| `COUNT(NA(),2)`    | 1         |
| `COUNTA(NA(),2)`   | 2         |

Referenced text and logical cells are excluded from numeric aggregates. COUNT excludes ordinary errors; COUNTA includes them and formula-returned empty text. Actual empty cells are excluded. Unsupported calculations, circular dependencies and resource-limit errors are still reported, including inside COUNT/COUNTA.

Supported named, table and INDIRECT references use the same provenance rules. INDEX can supply a cell, row or column reference to these aggregates, including zero row/column selectors; this does not implement dynamic-array spilling. Large ranges skip the empty tail outside used extents. Traversal and cell reads remain bounded, including sparse ranges with a distant used cell.

## Evidence

`tests/fixtures/aggregates-input.json` contains 128 formulas evaluated in native Excel 16.0 build 4627. `aggregates-oracle.json` records exact input SHA-256, formulas, environment, values and errors. `src/aggregate.test.ts` checks that binding, all native results, and unsupported/circular error propagation.

Two old unit expectations were corrected only after recording their exact native results: COUNTA includes the explicit empty-string argument, and AVERAGE of nonnumeric direct text returns VALUE. Large-empty-range tests now verify bounded skipping; a separate distant-used-cell case must still hit the work limit.

`tests/aggregates.spec.ts` imports a workbook with deliberately stale caches, checks mixed numeric/text/logical/error/empty cells, edits an error into a number, saves/reloads and inspects recalculated XLSX caches including a downstream dependent formula.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-formulas.ps1 -InputPath tests/fixtures/aggregates-input.json -OutputPath tests/fixtures/aggregates-oracle.json
npm.cmd run test -- src/aggregate.test.ts
npm.cmd run test:e2e -- tests/aggregates.spec.ts
```

## Limits and sources

Array constants, spill behavior, arbitrary reference-returning functions, locale-complete text coercion, exhaustive error precedence and exact floating-point edge behavior remain incomplete. This is tested aggregate behavior, not full Excel parity.

Primary sources: [Microsoft COUNT reference](https://support.microsoft.com/en-us/excel/functions/count-function), [SUM reference](https://support.microsoft.com/en-us/excel/functions/sum-function). The native receipt records the additional cases used to verify this implementation.
