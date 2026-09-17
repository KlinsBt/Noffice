# Calculated-table row resizing

Table design can now grow and shrink supported tables containing calculated columns. The range keeps the same header row and columns. The entire range/formula change participates in one undo and IndexedDB save transaction.

## Implemented behavior

- Growing a table fills blank cells in its calculated columns using the stored first-body-row master formula. Relative A1 references shift; absolute and structured references retain their meaning.
- Existing destination values/formulas and existing body exceptions remain intact. A blank or literal exception in the previous last row does not change which master fills the added rows. Explicit text-typed empty cells are preserved conservatively.
- Shrinking retains excluded worksheet cells. Local structured selectors in excluded formulas become table-qualified, including formulas in ordinary columns. Qualified current-row formulas outside the table return `#VALUE!`; regrowing restores their valid context. Ordinary A1 formulas continue to calculate outside the table.
- Removing the last nonblank body value clears that column's master in the same edit. Later growth does not restore deleted formulas. Undo restores both cells and master metadata.
- Retained/new XLSX export updates table/filter ranges, writes generated cell formulas, recalculates supported dependencies and preserves unrelated package content. Original file bytes remain separate.
- Legacy snapshots containing the old calculated-column resize restriction are rechecked against their original/checkpoint. Real source restrictions remain active.

## Architecture and limits

`sheet-tables.ts` shares range/style validation between the editor and retained writer; the writer validates without generating another set of cells from its older baseline. Formula population uses the complete proposed model before validating generated cells. `table-references.ts` qualifies excluded selectors without changing literals, quoted sheet names or already qualified tables. `clearRemovedColumnFormulas` runs on ordinary editor commits and preserves undo history.

Existing guards remain for protection, totals/hidden headers, active filters, sorting metadata, external data, pivots/slicers, table extensions and overlapping tables/merges. Worksheets with array/data-table formulas conservatively reject table resizing, including arrays outside the current table that could intersect a larger range. Excluded formulas with unsupported external references reject the operation atomically. Old and proposed ranges must fit the editable grid; the new rectangle is bounded to 10,000 rows, 256 columns and 100,000 cells.

Paste/horizontal expansion, moving the header, changing columns, table-aware row insertion/deletion, full formatting/validation inheritance and autofill preferences remain incomplete. The existing formula evaluator and validation limitations still apply. The resize command retains explicit cell formatting; it does not certify all differential-style inheritance or desktop rendering.

## Tests and native evidence

Fourteen unit cases in `src/table-resize.test.ts` cover the model, reference qualification, growth/shrinkage, clearing masters, exceptions, export/reimport, original package preservation, legacy metadata and rejection paths. Two browser workflows in `tests/table-resize.spec.ts` operate Table design, undo/redo, saved reload, actual downloads and reimport on retained and new workbooks. The existing calculated-column entry workflows also pass after the clearing change.

`scripts/verify-excel-table-resize.ps1` performs equivalent operations in an isolated native Excel 16.0 build 4627 instance. Five retained stages and two new-workbook stages compare 340 cell snapshots (values/formulas/number formats), table/header/filter ranges and master formulas. The retained stages also compare names, notes and validation/conditional rules. Final native comparison: zero mismatches, with all 15 source/browser/native-copy hashes matching. Full verification passes 1,499 unit tests (one opt-in skip), zero Svelte diagnostics, a 31-file static build, all 72 Chromium workflows and all 63 corpus scenarios. The strict corpus gate retains its two documented evidence gaps. [Checklist](../TASKS.md).

The verifier opens source and browser exports read-only with macros/events disabled, temporarily enables table formula autofill and restores the preference. It writes only authored comparison copies under `.local/xlsx-table-resize`, explicitly selects XLSX for new workbooks, and closes source edits without saving. Its JSON receipt binds source/browser/native-copy hashes to the results. It never modifies the supplied price schedule.

```powershell
npx.cmd vitest run src/table-resize.test.ts src/sheet-calculated-columns.test.ts
npm.cmd run check
npm.cmd run build
npx.cmd playwright test tests/table-resize.spec.ts tests/calculated-columns.spec.ts
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-table-resize.ps1
```

## Primary sources

Microsoft documents the explicit range workflow in [Resize a table](https://support.microsoft.com/en-us/excel/resize-a-table-by-adding-or-removing-rows-and-columns-in-excel) and column expansion in [Calculated columns](https://support.microsoft.com/en-us/excel/use-calculated-columns-in-a-table-in-excel-for-the-web). Native probes separately measured preservation of destination data and exceptions, first-row formula anchoring, cleared masters and excluded-row formula errors. These are scoped interoperability checks, not evidence of overall desktop parity.

Supported direct entry below tables and formula-triggered totals activation are covered by [the subsequent delivery](EXCEL_TABLE_ENTRY.md).
