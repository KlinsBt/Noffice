# Excel calculated columns

Direct formula entry in a table body now creates and updates calculated columns. The formula bar and in-cell editor use the same command. This is a bounded implementation of calculated-column entry, not certification of the complete Excel table feature set.

## Behavior

- Entering a formula into an otherwise empty table column fills every body row above and below the entry cell, including hidden rows. An existing value in the entry cell can be replaced when the other body cells are empty.
- Changing a uniform calculated column updates all its body cells. A1 relative references shift by row; absolute references and structured selectors retain their meaning. The master formula is stored relative to the first body row.
- Existing values elsewhere prevent automatic filling. A literal, blank, text-typed formula or different formula in an established calculated column is an exception: subsequent formula entry changes only the entered cell and retains the original master formula.
- Headings and totals are excluded. Styles, notes and other cell metadata survive the fill. All affected cells and table metadata share one undo/save transaction.
- XLSX export writes both cell formulas and the table's `calculatedColumnFormula`, updates supported dependent caches and preserves unrelated package parts. Header renaming repairs the master formula too. New workbooks and imported retained packages use the same table writer.
- Legacy snapshots recover missing master metadata from their original/checkpoint, preserving cell edits, exceptions and revision. Tables created after import receive an empty master array when no original part exists.
- Protected sheets/workbooks, overlapping merged cells, unsupported external/extended table metadata and intersecting array/data-table formulas reject automatic filling. Numeric Stop validation checks the proposed cells with row context before committing. Limits are 10,000 rows, 256 columns and 8,192 formula characters.

## Implementation

`sheet-calculated-columns.ts` performs the immutable entry command. `SheetTable.calculatedColumns` holds one formula string or null per column; strings use OOXML convention without a leading equals sign. `calculationBlocked` retains an import-time restriction. The importer also keeps the original package independently.

`SheetEditor.svelte` invokes the entry command only for direct formula entry. Explicit paste, fill, delete and search operations retain their own target-cell behavior; emptying the whole body removes its master. `xlsx-tables.ts` changes the individual formula child while preserving column IDs, attributes, totals and unrelated nodes. Supported [calculated-table row resizing](EXCEL_TABLE_RESIZE.md) now fills added blank cells and preserves excluded worksheet cells.

## Evidence and reproduction

The 24 cases in `src/sheet-calculated-columns.test.ts` cover propagation, exceptions, relative and structured references, hidden rows, totals exclusion, validation, preservation, backups, rename, legacy hydration, new exports, unsupported metadata and resource bounds. Two workflows in `tests/calculated-columns.spec.ts` exercise the actual static app, atomic undo/redo, formula-bar entry, saved reload, original/retained/new exports and reimport.

`scripts/verify-excel-calculated-columns.ps1` independently repeats seven retained-workbook stages and one new-workbook scenario in installed desktop Excel 16.0 build 4627. It compares 302 cell snapshots (values, formulas and number formats), table names/headings, hidden-row state and all eight master-formula arrays. The final comparison has zero mismatches; all 17 source/browser/native-copy hashes match the receipt in `.local/xlsx-calculated-columns/native-report.json`. Full verification passes 1,485 unit tests (one opt-in skip), zero Svelte diagnostics, a 31-file static build, all 70 browser workflows and all 63 external corpus scenarios. The strict corpus preservation gate retains its two documented gaps. [Checklist](../TASKS.md).

The verifier creates an isolated invisible Excel instance with macros/events disabled and temporarily enables calculated-column autofill, restoring the preference afterward. It opens source and browser outputs read-only, saves only authored native comparison copies under `.local`, closes without saving source edits, and never changes the supplied price schedule. It explicitly selects XLSX for the new workbook because the installed Office default is OpenDocument.

```powershell
npx.cmd vitest run src/sheet-calculated-columns.test.ts
npm.cmd run check
npm.cmd run build
npx.cmd playwright test tests/calculated-columns.spec.ts
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-calculated-columns.ps1
```

## Remaining scope

Automatic table expansion, row/column insertion, error-indicator/restore-column UI and per-user autofill preferences are incomplete. Supported row-range resizing is covered separately in [EXCEL_TABLE_RESIZE.md](EXCEL_TABLE_RESIZE.md). Formula comparison conservatively requires the stored translated expression to match each body formula exactly; mathematically equivalent expressions with different spelling are treated as exceptions. Comprehensive paste/autocorrect behavior, coercion, array/spill formulas and complete validation semantics are not certified. The existing evaluator's unsupported-function and cache limitations still apply. Totals-row creation and full table-aware sorting remain separate work.

## Primary sources and native findings

Microsoft documents automatic entry and subsequent column updates in [Use calculated columns in an Excel table](https://support.microsoft.com/en-us/excel/use-calculated-columns-in-an-excel-table). Its [formula-error guidance](https://support.microsoft.com/en-us/excel/detect-formula-errors-in-excel) describes exceptions to the master formula. Local native probes separately established the existing-value rule, relative-reference anchoring, hidden-row filling and preservation of exceptions. These measured results supplement the documentation; they do not establish general desktop parity.
