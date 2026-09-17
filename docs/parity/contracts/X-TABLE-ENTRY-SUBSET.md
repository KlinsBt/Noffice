# X-TABLE-ENTRY-SUBSET: direct entry below an eligible table

Status: partial relative to X-TABLE. Existing bounded behavior has passing tests; the entire table command family is not verified. Owns no complete ribbon command. Related workflows: direct grid/formula-bar entry, table design, undo, redo and XLSX export.

## Existing behavior and implementation

[EXCEL_TABLE_ENTRY.md](../../EXCEL_TABLE_ENTRY.md) defines eligibility, literal growth, calculated-column propagation, formula-triggered totals activation, stored labels/custom formulas, atomic validation and source retention. Source: src/sheet-table-entry.ts, sheet-tables.ts, xlsx-tables.ts and SheetEditor.svelte. Imported totals labels that look numeric are stored as text, matching native Excel. Unsupported structures retain ordinary entry behavior.

## Existing reproducible evidence

- Unit: npm.cmd run test -- src/sheet-table-entry.test.ts (22 cases recorded in the September 9 delivery).
- Browser: npx.cmd playwright test tests/table-entry.spec.ts; creates six .local/xlsx-table-entry exports, tests undo/redo/reload and retained/reimport behavior.
- Native: powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-table-entry.ps1; 384 cell snapshots plus types/formulas/format/table/totals/name/rule/note checks. Receipt explicitly normalizes absent totalsRowFunction to none and independently compares COM aggregation settings.
- Existing full regression evidence and original/native/browser hash audit: root TASKS.md and docs/TESTING.md. These are historical receipts, not a newly run full matrix for this planning change.

## Remaining acceptance work

- [ ] Specify and test actual native typing/UI behavior in addition to COM entry observations.
- [ ] Add paste, horizontal and Tab growth, preferences and formatting/validation inheritance contracts.
- [ ] Cover filtered/extended tables, preset totals, totals lifecycle and full table-column operations.
- [ ] Define stress/performance bounds and cross-browser interaction coverage.
- [ ] Bind all ten dimensions to current implementation/contract/baseline hashes before a full contract claim.

Handoff September 10: preserve this verified subset while following the model-foundation queue in ENGINE_PLAN.md. Do not assign the entire Table Design or totals command to this subset to make its checkbox green.
