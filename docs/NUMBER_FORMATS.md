# Excel number and date formats

The Format cells dialog is available from the Excel toolbar or Ctrl+1. It applies a preset or custom code to the selected range, previews the active cell, and participates in undo, local persistence and XLSX export. Formatting leaves the underlying cell values unchanged.

## Implemented subset

`numeric-format.ts` handles fixed decimal formats, required/optional/space digit placeholders, grouping, comma scaling, percentages, literal labels, currency symbols, positive/negative/zero/text sections, numeric conditions and bounded mixed/improper fractions. Decimal rounding uses 15 significant decimal digits with integer arithmetic. `number-format.ts` delegates other families to locally bundled SSF and adapts date separator dots, A/P letter case and single-digit fractional seconds. The same formatting entry point serves worksheet display and supported TEXT calculations.

Formatting accepts at most 2,048 code characters, 100 decimal places, 100 percent/scaling tokens, four variable denominator digits or a fixed denominator of at most 1,000,000. Exceeding those supported limits raises an explicit error. The dialog validates before applying.

## Independent evidence

- `tests/fixtures/number-formats-input.json` contains 44 independently authored numeric cases. The recorded native results in `number-formats-oracle.json` come from Microsoft Excel 16.0 build 4627, language 1031. A SHA-256 assertion binds the receipt to the exact input bytes.
- `number-format-edges-input.json` and its native oracle add 11 cases for omitted integer placeholders, rounded negative zero, conditional sign selection and large integers. These cases exposed errors beyond the downloaded corpus and drove further fixes. Both fixture receipts are hash checked: 55 native results in total.
- `scripts/verify-excel-number-formats.ps1` creates a separate native workbook and evaluates TEXT in Excel. It records native format strings, adjusts the authored 1904 date serials to equivalent 1900 serials, uses explicit decimal/grouping separators and restores application settings. Generated workbooks and fresh receipts stay in `.local/`.
- `src/number-format.test.ts` checks all 55 results, fixture provenance, literals and resource bounds: 59 tests.
- The downloaded Apache POI corpus contributes 325 numeric-format and 45 date-format stored results. All 370 match. These source-cache comparisons are distinct from the freshly calculated native oracle.
- `tests/number-formats.spec.ts` checks keyboard opening, live preview, range application, undo/redo, fraction formatting, invalid-code rejection, saved reload and actual XLSX downloads. Export assertions check underlying values, number-format codes and every unrelated original ZIP payload.

To regenerate native evidence on Windows with Excel installed:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-number-formats.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-number-formats.ps1 -InputPath tests/fixtures/number-format-edges-input.json -OutputPath .local/number-format-edges-oracle.json
```

Inspect the resulting receipt before replacing the committed oracle. Do not regenerate expected results from the application formatter itself. Hash-bound inputs are excluded from automatic Prettier rewriting; any intentional input change requires a new native receipt.

## Remaining limits

This does not implement every Excel format. Conditional format colors are retained in format codes but do not override the rendered font color. Width-dependent repeated fill and glyph-width alignment are incomplete (`*` contributes no fill and `_` contributes one space). Locale/calendar directives do not implement a complete locale engine. Scientific and date families retain SSF limitations. Format validation is bounded and sample-based, not a full Excel grammar validator. TEXT argument coercion is not fully compatible for all nonnumeric inputs. Native evidence covers textual results, not Excel's cell-width clipping, printed pages or every display locale.

Primary references: [Microsoft custom-format guidelines](https://support.microsoft.com/en-us/excel/review-guidelines-for-customizing-a-number-format), [TEXT](https://support.microsoft.com/en-us/excel/functions/text-function), and [date/time formatting](https://support.microsoft.com/en-us/excel/format-numbers-as-dates-or-times).
