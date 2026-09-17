# Scalar Excel statistics

The calculator implements 27 statistical/product function names, including legacy aliases:

- VAR, VAR.S, VARP, VAR.P; STDEV, STDEV.S, STDEVP, STDEV.P.
- MEDIAN, LARGE, SMALL; RANK, RANK.EQ, RANK.AVG.
- PERCENTILE, PERCENTILE.INC, PERCENTILE.EXC; QUARTILE, QUARTILE.INC, QUARTILE.EXC.
- PRODUCT, GEOMEAN, HARMEAN, DEVSQ, AVEDEV, MODE, MODE.SNGL.

`statistics.ts` owns the numeric algorithms. `formulas.ts` distinguishes direct arguments from references, names, tables and supported INDIRECT/INDEX results. Referenced text/booleans are excluded for ordinary aggregates; direct coercion and special MODE behavior follow the tested native cases. Variance uses centered differences to reduce cancellation with large offsets. Existing dependency tracking recalculates downstream cells and exported caches.

## Evidence

`tests/fixtures/statistics-input.json` contains 114 formulas evaluated by native Excel 16.0 build 4627. The corresponding `statistics-oracle.json` records results, errors, formula text, environment and the exact input SHA-256. `src/statistics.test.ts` verifies that binding and compares every result. Native numeric expectations use absolute tolerance 5e-10; errors must match exactly.

Cases cover sample/population variance, standard deviation, median, ties, ascending/descending ranking, percentile interpolation and excluded endpoints, truncated quartiles, empty/text/logical ranges, numeric string arguments, named constants/ranges, INDIRECT, INDEX, ordinary errors and large offsets.

The native comparison exposed behavior that should not be replaced by guesses: the tested SMALL fractional ranks truncate while LARGE rounds up, MODE ties follow the first encountered tied value, and scalar MODE text/logical arguments can return VALUE errors. Empty ranges have function-specific errors. These are recorded observations for the installed Excel version, not a claim about all future builds or all argument forms.

`tests/statistics.spec.ts` checks an actual workbook edit, dependent recalculation, Ctrl+Z/Ctrl+Y, save/reload and XLSX formula caches. The test deliberately starts with stale caches so a cached-value-only implementation cannot pass.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-formulas.ps1 -InputPath tests/fixtures/statistics-input.json -OutputPath tests/fixtures/statistics-oracle.json
npm.cmd run test -- src/statistics.test.ts
npm.cmd run test:e2e -- tests/statistics.spec.ts
```

The native helper creates a temporary workbook, disables macros, closes only its own workbook and restores its application settings. It does not overwrite user files. Excel rejected `RANK(2,"2")` at formula assignment; it is not presented as a calculated oracle result.

## Limits and sources

This is scalar calculation, not dynamic-array, distribution, regression or complete statistical parity. Array constants/spills, A-suffixed variants, locale-complete coercion, all reference-returning functions, exhaustive error precedence and extreme floating-point behavior remain open. The existing 100,000-cell calculation limit applies. Modern functions have compatibility with the tested Excel build, not a blanket Microsoft 365 certification.

Primary sources: [Microsoft STDEV.S documentation](https://support.microsoft.com/en-us/excel/functions/stdev-s-function), [statistical function inventory](https://support.microsoft.com/en-us/excel/statistical-functions-reference), [PERCENTILE.EXC interpolation](https://support.microsoft.com/sl-si/excel/functions/percentile-exc-function). Native receipts resolve the tested behaviors where prose is incomplete.
