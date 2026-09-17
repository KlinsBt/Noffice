# Excel cell hyperlinks

The spreadsheet editor supports inserting/editing a cell hyperlink with Ctrl+K or the Hyperlink command, editing its display text and ScreenTip, removing the selected link(s), undo/redo, saved reload, and following supported links. Open link explicitly opens web/email destinations or navigates within the workbook. Merely importing or selecting a cell does not follow its link.

Workbook navigation supports quoted sheet names (including escaped apostrophes), absolute cell references, rectangular ranges and supported defined names that resolve to ranges. Missing/hidden sheets, hidden destination cells and locations outside the loaded grid report an error. Web/email destinations are limited to HTTP, HTTPS and mailto. Imported unsupported destinations remain stored and preserved but cannot execute through Open link. External pages use a separate browsing context without an opener.

## Data and preservation

`sheet-links.ts` owns address checks, immutable editing and workbook-location resolution. `SheetLinkDialog.svelte` exposes labels, destinations and ScreenTips. Formula display text is read-only in this dialog. Adding a destination to a number preserves its numeric type; changing its display label explicitly creates text. New links receive a blue underline while other font properties remain. Removing a link retains cell contents and direct formatting. Protected-sheet link editing is refused; the full native protection permission matrix remains unfinished.

`xlsx-import.ts` bypasses ExcelJS hyperlink conversion and reads hyperlinks separately through `xlsx-links.ts`. This fixes numeric/formula cells being converted to text objects by the library. Source `location`, relationship destinations and ScreenTips are read from original XML. Missing single-cell link anchors within the supported grid are materialized. Range links attach to represented cells; empty cells in a large linked range are not all materialized.

The retained writer updates worksheet hyperlink nodes and only the necessary relationships. Editing/removing a member of a rectangular link splits its remaining range into rectangles; shared destinations remain linked. Unknown attributes, unrelated worksheet nodes, notes, page settings and other ZIP parts are preserved. Orphaned hyperlink relationships are removed without removing non-hyperlink relationships. New workbook exports patch the same hyperlink metadata after generating cell values, so linked formulas/numbers stay formulas/numbers. Unchanged exports still return the source bytes.

## Evidence

- Twelve tests in `src/sheet-links.test.ts` cover address rejection, original internal/range links, defined-name navigation, numeric/formula typing, protected editing, retained package updates, shared range splitting and new workbook export.
- `tests/sheet-links.spec.ts` imports a real XLSX, edits labels/addresses/ScreenTips, rejects unsafe addresses, tests undo/redo, navigates to an apostrophe-named worksheet, refuses protected edits, removes one range member, reloads, downloads/reimports and compares unrelated ZIP payloads. Its web popup uses an intercepted local response and asserts a null opener; no real destination request is sent during this test. Screenshot inspected.
- `scripts/verify-excel-links.ps1` opens the actual source and browser export read-only in installed Excel, with link updates, macros and events disabled. It verifies addresses, labels, ScreenTips, numeric/formula types/results, retained neighboring links, font, note and worksheet protection, then renders both PDFs. The SHA-256 receipt is `.local/xlsx-validation/links-report.json`.

```powershell
npx.cmd vitest run src/sheet-links.test.ts
npx.cmd playwright test tests/sheet-links.spec.ts
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-links.ps1
```

This is not complete Excel hyperlink parity. HYPERLINK formula evaluation, automatic URL recognition, file/network destinations, external workbook navigation, drawing hyperlinks, full defined-name expressions, full-range empty-cell materialization, hover/activation behavior and all browser/protection permutations remain incomplete. Native metadata assertions and PDF creation do not establish pixel-identical worksheet rendering.

Primary references: [Microsoft Excel link workflows](https://support.microsoft.com/en-us/excel/work-with-links-in-excel), [SpreadsheetML hyperlink attributes](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.hyperlink?view=openxml-3.0.1).
