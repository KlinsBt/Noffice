# Excel table and column renaming

Table design edits the table name and one selected column name. The column selector defaults to the active cell's column. Applying a rename repairs structured references before committing one undoable workbook change. Style/range changes entered in the same dialog join that change.

## Supported behavior

- Table names are checked against other tables, defined names and cell-reference syntax. The current table is excluded from collision checks. [International names](EXCEL_TABLE_HEADERS.md) support Unicode letters/backslashes as well as ASCII names, up to 255 characters.
- Column names accept nonempty text up to 255 characters. Duplicate names, ignoring case, are rejected. Apostrophes, brackets, `#`, `@` and punctuation are escaped/bracketed when inserted into references.
- Reference repair includes qualified selectors, column ranges, header/data/totals/current-row selectors, unqualified references within the owning table, worksheet formulas, defined names, calculated-column/totals formulas, standard chart formula elements and standard validation/conditional formulas.
- Literal strings stay literal. `INDIRECT("Sales[Quantity]")` therefore becomes `#REF!` after Sales is renamed, matching Excel. Supported formula caches and their ordinary dependents are recalculated; unsupported caches remain with an instruction for Excel to recalculate on opening.
- Table identities, column IDs, table styles, notes and unrelated package parts survive. Existing cell edits are exported into a working copy first. The private/imported original is never replaced.
- Imported workbooks use the existing `xlsxStructureBase` checkpoint for subsequent exports, undo/redo, IndexedDB and backups. New workbooks remain in model mode so adding sheets still works. Stale asynchronous results are rejected before committing; closing the dialog invalidates its pending result.

`workbook-table-rename.ts` validates the command and creates the checkpoint. `xlsx-table-rename.ts` traverses the relevant package parts and repairs references with the tokenizer in `table-references.ts`. This operates on source cells beyond the visible grid as well as editable model cells.

## Verification

The 22 cases in `src/table-rename.test.ts` cover token boundaries, escaped headings, unrelated tables/strings, ambiguous references, source preservation, consecutive checkpoints, formula/error caches, new workbooks and guarded dependencies. Chart and direct-rule cases are package-level tests; they do not establish chart rendering parity.

`tests/table-rename.spec.ts` imports the authored fixture, edits a data cell, rejects a duplicate column name, renames table/column, exercises undo/redo and saved reload, performs two further punctuation/escape renames, downloads each stage and reimports the final export. The inspected dialog screenshot is `test-results/excel-table-rename.png`.

`scripts/verify-excel-table-rename.ps1` opens the authored source and three actual browser exports in its own invisible desktop Excel instance. It performs equivalent renames and compares 144 cell snapshots (values, formulas and number formats), table/column names, defined names, validation/conditional formulas and the retained note. The final September 9 comparison reports zero mismatches, with all four source/export hashes verified. `.local/xlsx-table-rename/native-report.json` records the Office version/build and source/export SHA-256 hashes. No source is saved. TASKS.md records the full regression result.

```powershell
npx.cmd vitest run src/table-rename.test.ts
npm.cmd run check
npm.cmd run build
npx.cmd playwright test tests/table-rename.spec.ts
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-table-rename.ps1
```

## Remaining differences

- [Direct header-cell typing/paste](EXCEL_TABLE_HEADERS.md) now normalizes blank/duplicate headings and repairs references. Full coercion and localization remain open.
- Pivots, slicers, external links, queries/connections, data models, macros, protection, extension/alternate-content references and ambiguous unqualified references are guarded. A changed LET/LAMBDA expression is also guarded until scoped identifier repair is available.
- Exhaustive Unicode normalization equivalence, all external-reference forms, custom XML/query semantics and full chart cache/rendering refresh remain outside this verified subset.
- Hidden-header/totals formula metadata is repaired, but the full hide/show/totals user workflow is not validated here. Full calculated-column propagation remains separate work.
- OOXML whole-table references use `Table[#Data]`. Native Excel does not bind an unadorned name token in the fixture's XML to its table; renaming leaves those tokens alone. General normalization of bare table names during formula export remains open.
- Checkpoints remain limited to 50 MB, and this command inherits the importer/exporter's other limits. This is not desktop parity.

## Research and native findings

Microsoft documents automatic structured-reference repair when renaming tables/columns, the distinction between qualified and local references, and special-character escaping. [Structured-reference rules](https://support.microsoft.com/en-us/excel/using-structured-references-with-excel-tables). Table definitions retain separate name/displayName and column name metadata. [SpreadsheetML tables](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-tables).

Native Excel rejected the initial authored fixture's direct structured references in validation and conditional rules. The native fixture now uses a defined name for those rules; separate package tests retain direct-expression coverage. Native experiments also exposed the distinction between bare OOXML name tokens and table selectors. These are local experimental findings, not claims of complete behavior across Office versions.
