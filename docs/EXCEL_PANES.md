# Excel frozen panes

The Excel ribbon's View group now exposes Freeze at active cell, Freeze top row, Freeze first column and Unfreeze panes. Freeze at active cell fixes the rows above and columns to the left of that cell. These changes participate in undo/redo, local persistence and XLSX export. The controls reset the browser scroll position when the freeze area changes.

## Editing and preservation

`sheet-panes.ts` calculates the freeze boundary independently of the editor. It rejects a boundary that crosses a merged cell or leaves no scrollable row/column. Selecting A1 asks the user to select a boundary instead of calculating a viewport-dependent split. Worksheet protection does not prevent a view change; it continues to prevent edits to protected cells.

For imported workbooks, `xlsx-preserve.ts` patches the first represented worksheet view. It retains that view's zoom, gridline and unrelated attributes, other workbook windows, cell data, formulas, styles, notes and unrelated ZIP payloads. It replaces pane-related selections and creates a valid pane/active-cell selection for the new boundary. Unfreezing removes the pane. New workbooks continue to export through the existing ExcelJS view writer. Source originals remain available.

Frozen row headers, column headers and intersection cells use separate sticky offsets and drawing levels. Hidden rows and custom heights contribute to the visible offsets. `sheet-layout.ts` now keeps frozen rows mounted while virtualizing the scrolling body; it no longer renders every worksheet row merely because a freeze is active. The window grows with the viewport height and retains intersecting merge anchors.

## Evidence

- Seven tests in `src/sheet-panes.test.ts` cover commands and refusal paths, source identity/data retention, 10,000-row window bounds, short rows in a tall viewport, hidden rows, merged anchors, OOXML element ordering, preservation of extra views, freeze/unfreeze exports and reimported metadata.
- `tests/sheet-panes.spec.ts` imports a 2,000-row workbook with custom heights, a hidden row, a formula and a note. It freezes at C4, scrolls both axes, checks stationary headings and fewer than 65 mounted rows, then verifies undo/redo, saved reload, all four commands, downloads and reimport. Unrelated ZIP payloads remain byte-identical.
- `scripts/verify-excel-panes.ps1` opens the actual browser exports read-only in installed Excel 16.0 build 4627. The frozen file reports `FreezePanes=true`, `SplitRow=3`, `SplitColumn=2`; the unfrozen file reports false/zero/zero. The formula still evaluates to 3. The receipt in `.local/xlsx-validation/panes-report.json` binds each exported file's SHA-256.

```powershell
npx.cmd vitest run src/sheet-panes.test.ts src/xlsx-preserve.test.ts
npm.cmd run build
npx.cmd playwright test tests/sheet-panes.spec.ts
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-panes.ps1
```

The native helper creates its own Excel instance, disables macros/events, opens only generated test files and closes them without saving. It waits for Excel startup and retries transient busy responses. These tools and generated files are development-only; the application remains fully client-side.

## Remaining limits

Split panes with independent scrolling, editing multiple workbook windows, viewport-dependent A1 freezing, merge-crossing freeze boundaries and all native scrolling/selection behavior remain incomplete. Imported `frozenSplit` states are not fully modeled. A very large frozen area or merged range can still require many mounted rows; columns remain unvirtualized. The tests establish selected browser behavior and native pane metadata, not complete worksheet rendering or desktop Excel parity.

## Primary references

Research checked September 9, 2026:

- [Microsoft: Freeze panes to lock rows and columns](https://support.microsoft.com/en-US/Excel/get-started/freeze-panes-to-lock-rows-and-columns) describes the selection boundary and row/column/unfreeze commands.
- [Microsoft OOXML implementation notes: pane](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/bb178944-7838-4a58-9c16-e0c63d428a43) explains the frozen-pane split counts and top-left cell requirements. Native results above verify the generated files independently of Noffice's importer.
