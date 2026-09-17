# Excel data validation

The Data validation dialog creates or replaces Whole number, Decimal, List and Text length rules on selected cells. It supports eight numeric comparisons, Ignore blank, input messages/titles, Stop error messages/titles and Clear validation. Relative references start at the first selected cell. Existing values are retained when a rule is applied; future supported input is checked. Applying/clearing rules is one undo operation and persists through reload and XLSX export.

The dialog uses available desktop width for criteria and messages, with scrolling on narrow displays. New rules with unresolved list sources, unsupported/non-numeric bounds or reversed bounds are rejected. New rules are limited to 255-character formulas, native message/title lengths and 10,000 selected cells. Protected-sheet rule editing is refused. Formatting unchanged content does not revalidate its existing value.

## Model and import/export

`sheet-validation.ts` owns validation lookup, bounds, relative formula translation, definition checks and immutable commands. Optional `Sheet.validationRanges` retains original ranges without allocating every cell in a full-column rule. `Cell.validation` supplies explicit edits; `type: none` clears an inherited rule. The importer reads original SpreadsheetML rules independently of ExcelJS, preserving formulas rather than accepting its numeric conversion of formula bounds. Blank cells receive effective rules on selection/input.

`cell-validation.ts` resolves literal/named/range list choices and supported formula-based numeric bounds through the bounded formula evaluator. Input, paste and search replacement use the effective rule at the target address. The existing paste policy checks validation; this is stricter than some desktop Excel paste paths. Full native entry and alert semantics are not certified.

`xlsx-validation.ts` patches changed cells' rules in retained packages. It subtracts edited cells from source rectangles, clones the remaining ranges and translates relative formulas from the original anchor. Unrelated attributes and parts remain. New workbooks export original range metadata plus explicit cell overrides through the same writer. XML ordering and validation counts are updated. Extended x14 validation edits are rejected, avoiding silent replacement with standard rules. The operation has cell-count and processing limits.

## Evidence

- Fourteen cases in `src/sheet-validation.test.ts` cover lazy blank/full-column lookup, relative bounds, comparisons, protected edits, invalid setup, retained range splitting, unrelated payload equality and new-workbook export.
- `tests/sheet-validation.spec.ts` checks blank-cell enforcement, existing-value retention, formatting, applying and clearing rules, undo/redo, list choices, protection errors, saved reload, calculation, actual download/reimport and unrelated ZIP payloads.
- `scripts/verify-excel-validation.ps1` opens the actual source and browser export read-only in installed Excel. It verifies validation types/messages, the surviving relative formula, the large list range, unchanged notes/protection and values/formulas, and 14 independent numeric/list input decisions using native `Validation.Value`. Temporary in-memory input changes are restored before PDF rendering; source files are never saved. The hash-bound receipt is `.local/xlsx-validation/validation-report.json`.

```powershell
npx.cmd vitest run src/sheet-validation.test.ts src/cell-validation.test.ts
npx.cmd playwright test tests/sheet-validation.spec.ts
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-validation.ps1
```

## Remaining limits

Custom formulas, date/time creation and locale entry, Warning/Information confirmation flows, complete Ignore blank semantics, dropdown display flags, circles around invalid data, every named/dynamic source expression, all protection permissions, and x14 rule editing remain unfinished. Imported unsupported rules remain preserved, but unsupported evaluations currently do not reject input. Titles are stored/exported; the current message strip/notification is not a reproduction of native dialogs. Reopen older imports to populate range metadata. Complete Excel/Office desktop parity remains incomplete.

Sources: [Microsoft validation workflow](https://support.microsoft.com/en-US/Excel/get-started/apply-data-validation-to-cells), [existing values and alert behavior](https://support.microsoft.com/en-us/excel/more-on-data-validation), [SpreadsheetML validation schema](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datavalidation?view=openxml-2.20.0).
