# Excel table editing

This delivery adds table creation and a Table design dialog to the spreadsheet. It is a tested subset of Excel tables; it does not establish desktop parity.

## Implemented behavior

- Format as table creates a named table from a selected rectangle, using its first row as headers. Empty and duplicate headings receive unique names. Names are checked against workbook tables, defined names and cell references.
- Table design resizes the row extent while retaining the original header position and columns. Shrinking leaves worksheet values outside the new table intact. Structured-reference formulas and ordinary dependents recalculate.
- Six built-in styles, Medium 2–7, support banded rows/columns and first/last-column emphasis. None removes the table style. Imported style names remain selectable without replacement. Direct cell formats take precedence in the browser.
- New table filter ranges use the existing filter editor. Dropdown visibility metadata is distinct from active criteria and survives criteria changes. Clear criteria before resizing.
- Creation, style and range changes use the existing undo/redo, IndexedDB persistence and Noffice backup paths. New workbooks retain sheet creation. Imported XLSX exports preserve source parts and patch the affected table definitions, relationships, content types, cells and formula caches.
- The original XLSX remains separate and unchanged. [Table and column renaming](EXCEL_TABLE_RENAMING.md) is available in Table design with package reference repair. [Direct header typing/paste](EXCEL_TABLE_HEADERS.md) also repairs references and normalizes duplicate/blank names.

## Implementation

`src/sheet-tables.ts` owns range/name checks, commands, header guards and table appearance. `src/xlsx-tables.ts` patches existing table definitions or adds new parts. `SheetTableDialog.svelte` exposes commands without widening the Home toolbar. Original table relationships are read directly from OOXML, bypassing ExcelJS's assumptions about `tableN.xml` filenames and relative relationship targets.

Older saved workbooks recover additive table metadata from their retained source through `workbook-tables.ts`. Hydration keeps cell edits and revision numbers, ignores stale asynchronous results and does not queue a storage write. The preservation writer also restores metadata when exporting an older snapshot directly, including byte-identical unchanged export. Workbooks saved before original worksheet-path preservation still require reopening their original file.

Native inspection exposed two generator issues: ExcelJS emits tableParts before legacyDrawing in the combined table/note fixture, and marks its Office 2007 theme with a newer defaultThemeVersion. The authored fixture corrects the element order and removes the mismatched default marker. New table workbook exports also remove that generator marker so Excel uses the embedded theme. They inherit the workbook font until a user supplies direct font formatting; synthesizing a direct font otherwise masks table header colors. Imported originals are never rewritten to work around the theme marker.

Table band colors use Excel's integer HLS conversion and luminance truncation, verified against actual displayed colors. Theme substitution by different Office versions remains outside the validated subset; a source whose defaultThemeVersion contradicts its embedded theme can still display differently. Custom table styles, differential formatting overrides, complete borders and totals-row style rendering are not implemented.

## Verification

- `src/sheet-tables.test.ts`: 25 cases covering creation, invalid ranges/names, collisions, protection, overlaps, header guards, resize/cache updates, themes, button metadata, filtered new exports, Noffice backup, older snapshot metadata and preservation of unrelated XML.
- `tests/sheet-tables.spec.ts`: three static-build workflows covering imported/new tables, six styles, undo/redo, filter guards, saved reload, download/reimport, duplicate headers, sheet creation and reopening an older IndexedDB record with existing edits.
- `scripts/verify-excel-tables.ps1`: opens the source and eleven actual browser exports in native Excel, applies equivalent operations and compares 462 cell snapshots, table definitions/options and notes. It also checks the new-workbook export and compares 33 browser-rendered style cells, including column bands, first/last-column emphasis and None. Receipts and SHA-256 hashes are written under `.local/xlsx-tables/`.
- UI screenshot: `test-results/excel-table-design.png`. It shows the range/name fields side by side and paired style options.

The task checklist records the final passing status. Failed intermediate native comparisons are diagnostic evidence, not acceptance.

## Reproduction

```powershell
npm.cmd run check
npx.cmd vitest run src/sheet-tables.test.ts
npm.cmd run build
npx.cmd playwright test tests/sheet-tables.spec.ts
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-tables.ps1
```

The native verifier opens generated files read-only in its own invisible Excel instance, disables macros/events, closes without saving and never edits the private original price schedule.

## Open scope

- Complete table/column renaming semantics beyond the [verified dialog/package subset](EXCEL_TABLE_RENAMING.md) and [direct header entry](EXCEL_TABLE_HEADERS.md), including external/scoped dependencies.
- Column insertion/removal, paste/horizontal expansion, conversion to ranges and table removal. Direct [calculated-column entry](EXCEL_CALCULATED_COLUMNS.md) and [supported row resizing](EXCEL_TABLE_RESIZE.md) are implemented; complete paste/autocorrect and restore-column behavior remain open.
- Header-row insertion/hiding, totals-row commands/functions and empty-data-body semantics.
- Resize of tables with totals, hidden headers, external data, pivot/slicer dependencies, sort metadata or extension metadata, or worksheets with array/data-table formulas. Protected workbooks/sheets reject edits. These guards preserve source packages.
- Custom table styles, all built-in styles, complete theme substitution, differential-format precedence, borders and exact totals-row appearance.
- Creation/resizing beyond 10,000 rows, 256 columns or 100,000 cells. Imported range metadata is read without eagerly materializing the full rectangle.

## Primary references

- [Microsoft: create and format tables](https://support.microsoft.com/en-us/excel/get-started/create-and-format-tables)
- [Microsoft: resize a table](https://support.microsoft.com/en-us/excel/resize-a-table-by-adding-or-removing-rows-and-columns-in-excel)
- [Microsoft: table ranges and resize semantics](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-tables)
- [Microsoft Open XML: table definitions](https://github.com/OfficeDev/open-xml-docs/blob/main/docs/spreadsheet/working-with-tables.md)
- [Microsoft Open XML: table style metadata](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyleinfo?view=openxml-3.0.1)
- [Microsoft Open XML: default theme version](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.workbookproperties.defaultthemeversion?view=openxml-3.0.1)

The generator corrections and integer color behavior above are findings from local source inspection and native Excel experiments, rather than claims supplied by those API reference pages.

Supported direct entry below tables and formula-triggered totals activation are covered by [the subsequent delivery](EXCEL_TABLE_ENTRY.md).
