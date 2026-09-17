# Automatic table expansion from direct entry

Typing a nonempty value into the first worksheet row below an eligible table now extends that table by one row. In-cell editing and the formula bar share the command, and the value, range, filter range and generated formulas form one undo/save transaction.

The command accepts entry under any table column. Other cells in that row must be empty; existing neighboring data prevents expansion. Replacing an existing value in the entered cell can expand the table. Gaps, empty input and entry outside the table's columns leave the range unchanged.

Literal entry fills blank calculated cells from their stored master formulas. Formula entry activates a totals row, restores supported stored labels/custom formulas, and excludes that row from structured data sums. This distinction was measured in native Excel; treating it as another data row produced incorrect dependent totals. Literal entry in a calculated column remains an exception. Ordinary edits and deletion of visible totals synchronize their labels/custom formulas for export. Existing body formulas and column masters are preserved.

`sheet-table-entry.ts` first checks the resize constraints without generating cells, applies the entered value immutably, and uses `editSheetTable` to change the range and generate eligible formulas. Validation failures reject the whole automatic change. Unsupported resize structures fall back to ordinary worksheet entry; existing protection and validation rules still govern that entry.

## Limits

This delivery covers direct entry only. Paste-triggered growth, horizontal expansion, Tab-to-add-row behavior, autofill preferences, full format/validation inheritance and expansion of filtered/already-totalled/extended tables remain incomplete. Preset totals-function activation, totals toggling/movement and the totals function dropdown remain incomplete; imported preset/extended totals settings conservatively disable implicit totals activation. Unedited source settings are retained. Existing resize bounds and guards still apply: 10,000 rows, 256 columns, 100,000 cells, table/merge overlaps, protection, arrays, external data and unsupported table dependencies. Explicit text-typed empty neighbors conservatively prevent expansion. These scoped checks do not establish overall Office parity.

## Verification

Twenty-two tests in `src/sheet-table-entry.test.ts` cover literals/zero, formulas, exceptions, occupied neighbors, gaps, constraints, validation, immutability and generated/retained XLSX output, totals/data separation, legacy hydration, labels/formulas/deletion and preset guards. `tests/table-entry.spec.ts` exercises the real static app through undo/redo, reload, formula-bar entry, six exports, retained ZIP checks and reimport.

`scripts/verify-excel-table-entry.ps1` repeats the six retained-workbook stages in its own invisible native Excel instance, comparing 384 cell snapshots plus table/filter/master/totals/name/rule/note metadata. Totals checks include visibility, custom formulas, labels, native COM aggregation settings and cell value types. The receipt explicitly normalizes an absent `totalsRowFunction` attribute to `none`; the native COM setting is compared independently. Numeric-looking totals labels are stored as text, matching native Excel. The occupied-neighbor case explicitly seeds existing data with expansion disabled before restoring it and entering the target cell. The verifier saves authored native copies under `.local/xlsx-table-entry`, opens source/browser files read-only with macros/events disabled, and restores both Excel autofill preferences. Source and browser originals are never overwritten. Final evidence and hashes are recorded in [TASKS.md](../TASKS.md).

```powershell
npx.cmd vitest run src/sheet-table-entry.test.ts
npm.cmd run check
npm.cmd run build
npx.cmd playwright test tests/table-entry.spec.ts
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-table-entry.ps1
```

Microsoft describes direct-entry expansion in [Resize a table](https://support.microsoft.com/en-us/excel/resize-a-table-by-adding-or-removing-rows-and-columns-in-excel). Microsoft also documents the distinct [Total Row](https://support.microsoft.com/en-us/excel/get-started/total-the-data-in-an-excel-table) and its function choices. Local native probes separately established the occupied-neighbor restriction and the distinction between literal and formula entry.
