# Excel subtotals, filtering and row visibility

Excel mode supports SUBTOTAL's eleven aggregate operations and their 101–111 hidden-row variants. Sheet settings can hide selected rows or unhide all rows on an unprotected, unfiltered sheet. The ribbon offers Filter range, Reapply filters and Clear filters. Filter range uses a selected header/data range or the existing worksheet/table filter containing the selected cell.

## Implemented behavior

- SUBTOTAL handles average, numeric count, nonempty count, maximum, minimum, product, sample/population standard deviation, sum and sample/population variance. It ignores nested SUBTOTAL expressions and recognizes AGGREGATE calls for nested exclusion; AGGREGATE calculation itself is not implemented.
- Tested arguments include multiple references, absolute names, contiguous table selectors, INDIRECT and bounded whole-column traversal. Referenced text and booleans do not become numeric inputs. COUNTA counts formula-produced empty strings and ordinary errors; COUNT ignores ordinary errors. Unsupported calculations remain explicit.
- Imported hidden rows, worksheet filter mode and table filter ranges determine visibility. Native Excel 16.0 tests show that active worksheet filter mode excludes hidden rows for code 9 as well as 109, including manually hidden rows outside the filter range. The receipt records this observed behavior separately from Microsoft's general manual-hide rules. Imported table filtering can also exist without worksheet filter mode.
- Changing cell values recalculates totals without implicitly reapplying filters. Reapply is an explicit operation. Filtering supports selected displayed values, blanks, scalar comparisons and wildcard text criteria; imported pairs of custom criteria retain AND/OR semantics. Multiple columns combine conjunctively.
- Existing worksheet/table filter criteria can be changed or cleared while preserving their source ranges and unrelated parts. Newly created workbooks export supported worksheet filters. Imported range resizing/removal and arbitrary table creation remain unsupported.
- Visibility/filter changes invalidate subtotal export caches and ordinary downstream dependents. Original unsupported caches stay retained. Range operations participate in undo/redo and IndexedDB saves.

## Evidence

`tests/fixtures/subtotal-{manual,filtered,table}-input.json` contains 44 cases per scenario, with **132 recorded native Excel results**. `filter-{numeric,wildcard}-input.json` adds six native subtotal results plus independent row-visibility observations. All receipts record Excel 16.0 build 4627, language 1031, exact fixture SHA-256, actual formula text and typed results. The filtered scenarios edit criteria cells after filtering to verify that visibility does not automatically change. The harness rejects invalid formulas explicitly rather than inventing an Excel result.

Regenerate a receipt using its matching input:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-formulas.ps1 -InputPath tests/fixtures/subtotal-filtered-input.json -OutputPath .local/subtotal-filtered-oracle.json
```

`src/subtotal.test.ts`, `src/sheet-filters.test.ts`, `src/subtotal-export.test.ts` and `src/filter-export.test.ts` cover native results, bounds, filter metadata, protected/unsupported operations, saved model validation, retained ZIP payloads and new/worksheet/table XLSX exports. The separate downloaded corpus assertion now checks all **352** formulas in `ConditionalFormattingSamples.xlsx`, including all 18 formerly unsupported subtotals.

`tests/subtotal.spec.ts` exercises row hiding, both subtotal variants, undo/redo, saved reload and downloaded caches. `tests/filters.spec.ts` exercises value/comparison filters, edits before reapply, undo/redo, saved reload, downloads and clearing.

`scripts/verify-excel-filter-export.ps1` opens the actual browser download read-only in a separate Excel instance. Fresh recalculation produces 2; reapplying exported criteria still produces 2; clearing criteria in memory produces 12. `.local/filter-validation/desktop-report.json` binds these results to the exported XLSX hash. The source download is not modified.

## Remaining limits

Color/icon/date-group/dynamic/top-ten/extension filter criteria are retained and their existing hidden-row state is usable by SUBTOTAL. They cannot yet be reapplied or edited through this dialog; clearing criteria is supported. Locale-sensitive labels, calendar grouping, merged headers, advanced filter dialogs and slicers remain incomplete. Changing a filter column currently reapplies all modeled filters, so an unsupported predicate elsewhere can block reapplication. Full-column filtering and large-sheet worker execution remain open.

The dialog displays up to 200 matching unique values at once; search narrows the list. Filtering and subtotal traversal are bounded. Full statistical precision, all reference-returning functions, three-dimensional reference syntax, outline/group creation and the desktop automatic subtotal insertion command remain open. Passing these cases does not establish full Excel or suite parity.

Primary references: [Microsoft SUBTOTAL](https://support.microsoft.com/en-us/excel/subtotal-function), [AutoFilter](https://support.microsoft.com/en-us/excel/use-autofilter-to-filter-your-data), [SpreadsheetML tables](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-tables), and [hidden-row representation](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/how-to-retrieve-a-list-of-the-hidden-rows-or-columns-in-a-spreadsheet).
