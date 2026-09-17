# Direct Excel table headers and international names

The spreadsheet accepts table-heading edits in the cell/formula bar, Delete and rectangular text paste. It repairs column references before committing a single undoable change. Table design remains available for explicit renaming.

## Header behavior

- Header names are normalized across the whole table, from left to right. The first duplicate keeps its spelling. Later duplicates receive a numeric suffix starting at 2, avoiding names already present anywhere in the supplied row. For example, `Price, Price, Price2` becomes `Price, Price3, Price2`.
- Blank headings receive sequential `Column1`, `Column2`, etc., then participate in duplicate normalization. A single cleared heading starts at Column1 regardless of its column position. Names are truncated to 255 characters before collision handling.
- A duplicate edit can rename another existing column. Both changes repair references together. Swapping headings similarly uses a simultaneous mapping rather than a sequence that could redirect formulas to the wrong column.
- Pasted body cells and header names are one transaction. If a protected or unsupported dependency prevents repair, no cell in the proposed rectangle is committed. Existing originals remain untouched.
- Apostrophe-prefixed literal headings are accepted, including `'=Heading`. Live formula entry in a heading is rejected with an instruction to enter text; general formula-to-header coercion is not certified.
- Pending reference repair makes the spreadsheet temporarily inert, exposes progress and disables Export. The unload guard includes pending header work. A stale result or unmounted editor cannot overwrite the active workbook. Normal editing resumes after success or rejection.

`sheet-header-entry.ts` prepares the proposed cells and normalized column commands. `workbook-table-rename.ts`, `xlsx-table-rename.ts` and `table-references.ts` support simultaneous column mappings. Unchanged pasted headings retain their text type. Header edits preserve the table's separate internal name when only columns change.

## International table names

Creation and Table design accept Unicode letters and backslashes as well as existing ASCII names. Calculation and formula copying recognize those identifiers, preserving names such as `銷售A1`, `Ümsatz.A1` and `\A1` while shifting actual A1 references. Case-insensitive collisions, spaces, leading digits and ambiguous R1C1 names are rejected. This is not an exhaustive Unicode normalization or locale-conformance claim.

## Evidence

- `src/sheet-header-entry.test.ts`: 17 cases for normalization, reserved suffixes, blank/long labels, simultaneous swaps, formula/error caches, unchanged pasted headings, retained ZIP parts, backups, new workbooks and atomic rejection.
- `src/table-names.test.ts`: 16 cases for Unicode/backslash calculation and export, copy boundaries, collisions, invalid names and preservation of an internal table name distinct from its display name.
- `tests/header-entry.spec.ts`: direct entry, Delete, paste with body cells, undo/redo, reload, nine downloads/reimport, unsupported-dependency rollback and delayed-work/export gating.
- `tests/table-names.spec.ts`: four consecutive international-name renames with calculation, saved reload, exports and reimport.
- `scripts/verify-excel-header-entry.ps1`: native comparison of nine actual browser exports, 432 cell snapshots (values/formulas/formats), column names, defined names, validation/conditional formulas and notes.
- `scripts/verify-excel-table-names.ps1`: four actual browser exports compared against native Excel, 192 cell snapshots and table/defined names.

The header verifier records four automatic-heading translations from the installed German Office's `SpalteN` to the app's English `ColumnN` before comparison. This locale mapping is explicit in its receipt; it is not a byte-for-byte claim across localized UI languages. Both verifiers open generated files read-only in their own invisible Excel instances, disable macros/events and close without saving source edits. Receipts under `.local/xlsx-header-entry/` and `.local/xlsx-table-names/` bind results to source/export SHA-256 hashes. TASKS.md records final acceptance.

The final September 9 run has zero mismatches across all 624 new native cell snapshots and the checked metadata; all 15 source/export hashes match. The earlier three rename exports also pass their 144-snapshot native regression. Full verification passes 1,461 unit tests, 68 Chromium workflows, 63 corpus scenarios, zero Svelte diagnostics and a 31-file static build. The strict corpus gate still reports the two established nested-DOCX/fully locked XLSX evidence gaps.

```powershell
npx.cmd vitest run src/sheet-header-entry.test.ts src/table-names.test.ts
npm.cmd run check
npm.cmd run build
npx.cmd playwright test tests/header-entry.spec.ts tests/table-names.spec.ts
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-header-entry.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-table-names.ps1
```

## Remaining scope

Full header coercion of formulas/dates/custom-formatted values, localized generated names, Unicode normalization equivalence, input during long operations, and commands for hidden headers/totals remain incomplete. Existing pivot/query/external/scoped/extension reference guards and the 50 MB checkpoint bound still apply. Direct [calculated-column entry and exception preservation](EXCEL_CALCULATED_COLUMNS.md) are now implemented. Automatic table expansion, full table-aware sorting, column insertion/removal and complete chart refresh remain separate outstanding work.

Microsoft documents direct heading edits in [Set up your header row](https://support.microsoft.com/en-us/excel/get-started/set-up-your-header-row), default headings in [Turn table headers on or off](https://support.microsoft.com/en-us/excel/turn-excel-table-headers-on-or-off), and reference/name rules in [Using structured references](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables). Duplicate reservation, simultaneous reference identity and the additional accepted backslash forms above were measured locally in desktop Excel.
