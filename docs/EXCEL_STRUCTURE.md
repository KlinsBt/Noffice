# Excel row and column operations

The Cells ribbon group's Rows and columns menu inserts or deletes whole rows and columns through the selected rectangular range. Selection width or height determines how many columns or rows are affected. Commands include undo/redo, saved reload and retained XLSX export. They work in new workbooks and the supported imported worksheet subset; they do not establish desktop Excel parity.

## Implementation

`sheet-structure.ts` repairs A1 references when worksheet coordinates change. Absolute and mixed references move with their cells; ranges expand or shrink; completely deleted references become `#REF!`. It distinguishes strings, function names, defined names and quoted sheet names. Entire-row/column and cross-sheet references are supported. Shared formula groups expand before repair. The calculator recognizes qualified errors such as `Data!#REF!`.

`xlsx-structure.ts` transforms a copy of the existing ZIP package. It moves rows, cell addresses, represented formatting and column spans; updates merge ranges, validation, filter/sort ranges, names, print areas/titles, supported selections, frozen panes and page-break references. Validation inherits from the preceding row/column at range boundaries. Hyperlink cell anchors move while their destination strings remain unchanged, matching native Excel. Formula caches and the calculation chain are invalidated and Excel is asked to recalculate. Styles, themes and unrelated package payloads remain intact.

`workbook-structure.ts` exports pending cell edits before transforming the package, then imports that copy as a new baseline. `WorkbookContent.xlsxStructureBase` stores a retained workbook baseline as base64 inside the same undo/persistence snapshot as the grid. This preserves earlier edits and lazy blank-cell rules through later structural operations. `OfficeFile.original` remains unchanged and downloadable. Subsequent edits use the preservation writer against the checkpoint. Undo restores both the package and grid. Stale or unmounted editors cannot commit a finished background operation into another document.

## Evidence and reproduction

- `src/sheet-structure.test.ts`: reference boundaries, quoted names, literals, shared formulas, validation inheritance, new workbooks, merged ranges, guarded features, checkpoint/backup reload and preserved payloads.
- `tests/sheet-structure.spec.ts`: actual browser commands, undo/redo, reload, eight downloaded packages, cross-sheet results, reimport and byte-identical rejection paths. Screenshot: `test-results/excel-structure.png`.
- `scripts/verify-excel-structure.ps1`: opens the browser source read-only in desktop Excel, performs equivalent native operations in memory, and compares each downloaded stage. Checks 1,008 cell snapshots across eight stages: formulas, calculated values, number formats, bold, fill and merged ranges. Also compares names, hyperlink locations, row heights, column widths, hidden states, validation and print settings. Receipt: `.local/xlsx-structure/native-report.json`, with source/export SHA-256 hashes and Excel version/build. Native workbooks close without saving.

```powershell
npm.cmd test -- src/sheet-structure.test.ts
npm.cmd run check
npm.cmd run build
npx.cmd playwright test tests/sheet-structure.spec.ts
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-structure.ps1
```

## Remaining scope

- Protected target worksheets, drawings, legacy notes, controls, custom views and unhandled anchors reject before content changes. Workbooks with tables, charts, pivots, external-link packages or threaded comments are also guarded. Ordinary supported cell editing remains available separately.
- Array/data-table formulas, structured/external/3-D references, separately sheet-qualified range endpoints, ambiguous unqualified workbook names, split panes and worksheet extensions need further repair support.
- Partial-cell shifts, insert-copy choice dialogs, full table restructuring, complete outline/group inheritance, cut/move repair and sheet rename/add/delete repair remain open. New workbooks stay in model mode so their existing sheet-creation workflow remains available.
- Conditional-format range/formula movement is implemented, but native rule-priority and insertion-inheritance coverage is incomplete. Frozen-pane/page-break repair does not establish complete desktop view/printing equivalence.
- The grid still has 10,000 rows and 256 columns. Structural work is bounded to 100,000 represented/inserted cells per target sheet and a 50 MB compressed checkpoint, with existing ZIP expansion limits. Checkpoints increase local storage and undo memory. Large Noffice backups remain subject to import limits.
- Unsupported calculations lose stale caches during structural repair and require recalculation by a supporting application. Missing formula functions and browser calculation limits remain.

## Primary sources

Microsoft documents the commands in [Insert or delete rows and columns](https://support.microsoft.com/en-gb/excel/get-started/insert-or-delete-rows-and-columns-in-excel), and deleted explicit references in [How to correct a #REF! error](https://support.microsoft.com/en-US/Excel/how-to-correct-a-ref-error). Package structure follows [Working with formulas](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas) and [Working with sheets](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-sheets). Exact inheritance and hyperlink behavior above were measured in installed desktop Excel; documentation alone was not treated as compatibility evidence.
