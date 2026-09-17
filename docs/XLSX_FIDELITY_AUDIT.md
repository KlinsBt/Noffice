# XLSX fidelity audit — 2026-09-08

**Result: failed workbook fidelity.** The supplied exported workbook differs substantially from its original in layout, retained features, data types, and cached formula results. The existing small self-generated round-trip tests did not establish fidelity for this workbook.

Compared `01 Price schedule (1) (1).xlsx` (exported) against `01 Price schedule (1).xlsx` (original). Both source files were read without modification. Their contents were treated as data; no workbook instructions, formulas, hyperlinks or external connections were executed. This is an OOXML and ExcelJS model inspection, not a Windows Excel rendering test. The private workbooks and detailed cell-value reports are not added to the public source tree.

## Measured losses

| Feature                                 | Original                           | Exported               |
| --------------------------------------- | ---------------------------------- | ---------------------- |
| File size                               | 98,422 bytes                       | 15,868 bytes           |
| Worksheets                              | 5, including 3 hidden              | 5, all visible         |
| Merged ranges                           | 25                                 | 0                      |
| Rows with explicit height               | 123                                | 0                      |
| Hidden rows                             | 61                                 | 0                      |
| Column widths                           | Sheet-specific widths              | A–Z all set to 18      |
| Border definitions                      | 22                                 | 1 empty/default border |
| Cell formatting records (`cellXfs`)     | 118                                | 26                     |
| Defined names, including print settings | 20                                 | 0                      |
| Worksheet protection                    | 4 sheets                           | None                   |
| Validation entries expanded by ExcelJS  | 102                                | 0                      |
| Image placements                        | 3 placements of one embedded image | None                   |
| Comment XML parts                       | 2                                  | 0                      |
| Structured table parts                  | 1                                  | 0                      |
| External-link definition parts          | 5                                  | 0                      |
| Formula cells                           | 187                                | 187                    |

Styles are deduplicated during export, so counts alone are not fidelity metrics. Direct inspection additionally confirms lost borders, text wrapping, vertical alignment, theme-based fills and colors, and rich-text runs. For example, the experts-sheet heading uses a theme fill and contrasting theme text in the original; both are absent in the export. Missing styled blank cells also remove the form's blank input areas and visual structure.

The main worksheet originally uses A4 landscape, a specific print area, repeated header rows, custom margins, page breaks and footers. The export switches to portrait defaults and loses the print area, repeated headers and footers. The experts sheet's frozen top ten rows and saved view settings are also lost.

## Data and calculation differences

ExcelJS was used to resolve shared-formula representations before comparing formulas. **All 187 formula expressions match** after that normalization; raw XML shared-formula serialization differences are not counted as lost formulas.

There are 21 changed cached formula results under ExcelJS's decoded representation. These include unsupported `VLOOKUP` calculations wrapped in `IFERROR`: existing lookup results become the fallback `N.N.`. Date-typed values and a date-returning formula become ISO date strings. Numeric-looking identifier text is converted to a number. Hyperlink values lose their link targets, and rich text becomes plain text. These are functional losses as well as appearance changes. Not every cached-result change independently proves a calculation defect; the unsupported-lookup path and date/type conversions are confirmed in the application code.

## Why the previous code caused this

1. `src/formats.ts` imports a narrow sparse cell model. It skips merged follower cells and does not store merge ranges, row heights, hidden state, column geometry, borders, wrapping, complete number formats, validation or sheet protection.
2. Theme/indexed colors are ignored unless a color has direct ARGB representation. Rich runs and hyperlinks are reduced to plain text. Dates become strings.
3. Export constructs a new `ExcelJS.Workbook`, creates every worksheet with default visibility and geometry, and explicitly assigns every column width to 18. It writes a small subset of cell properties and discards unsupported package parts.
4. Export recalculates formulas through Noffice's partial evaluator. Unsupported VLOOKUP produces an internal error that IFERROR turns into an apparently ordinary fallback string.
5. Retaining the original for a separate download does not preserve those features in the edited export.

## Acceptance criteria for the fix

These criteria were established against the failed export. The remediation below and **TASKS.md** distinguish passing implementation checks from the remaining Windows Excel fidelity gate. A file that opens successfully is not sufficient evidence.

- Preserve original OOXML parts and relationships for imported workbooks; patch supported edits without dropping unrelated features. An unchanged round trip must preserve the original workbook's supported semantics and retained parts.
- Keep cell data types, full number formats, rich runs, theme/indexed colors, cached results, and formula metadata. Distinguish an unsupported calculation from a genuine Excel error before writing a replacement result.
- Model and render widths, heights, merged cells, borders, wrapping, hidden sheets/rows, protection and freeze panes. Preserving export data alone does not fix the editor view.
- Preserve validation, names, tables, links, drawings, comments, print setup and other untouched workbook parts.
- Build independently authored, redistributable synthetic fixtures matching these structures; include unchanged and edited round trips and explicit no-loss assertions. Do not publish the supplied private workbook as an MIT fixture.
- Compare results in Windows Excel, including screen and print appearance, before checking the workbook-fidelity gate.

## Reproduction

`scripts/compare-xlsx.py` accepts original and exported paths and writes a detailed local JSON report. It compares package structure and raw cell representations; shared formulas and cached numeric/date representations require semantic normalization before interpreting differences.

```powershell
python scripts/compare-xlsx.py "path/to/original.xlsx" "path/to/exported.xlsx" --report .local/workbook-comparison.json
```

## Implemented remediation

The initial audit only diagnosed the failure. Subsequent implementation replaces imported-XLSX reconstruction with source-package patching in `src/xlsx-preserve.ts`. The supplied bad export and original files are unchanged; the repair applies to newly imported originals and subsequent exports.

Unchanged browser export of the supplied original is byte-identical. The edited browser scenario changes D27, verifies its F27 quantity-times-rate result, and compares the complete package entry list and uncompressed payloads of all non-worksheet parts except workbook recalculation metadata. This retains the original hidden sheets, styles, images/drawings, comments, validations, names, tables, external-link records and print settings. Formula-expression edits additionally remove obsolete optional calculation-chain metadata; ordinary input edits retain it.

The local comparison report for the repaired edited export records:

| Measurement                    | Original | Repaired edited export |
| ------------------------------ | -------- | ---------------------- |
| Compressed bytes               | 98,422   | 93,712                 |
| Package parts                  | 56       | 56                     |
| Worksheets / hidden worksheets | 5 / 3    | 5 / 3                  |
| Merged ranges                  | 25       | 25                     |
| Defined names                  | 20       | 20                     |
| Removed package parts          | —        | 0                      |

The compressed size changes because the edited package is recompressed. The six changed cell representations consist of the intentionally changed input and five dependent cached totals, including cross-sheet totals. Their formula expressions and shared-formula attributes remain unchanged. This comparison is recorded in ignored `.local/repaired-workbook-comparison.json`; the review file is `test-results/private-price-schedule-edited.xlsx`. These private artifacts are excluded from distribution.

The editor now interprets original dimensions, hidden state, merges, borders, styled blank cells, wrapping, vertical alignment, theme colors, native types, number formats and frozen panes. It displays raster drawings and cell notes, offers supported local named validation lists, and guards locked cells on protected sheets. VLOOKUP results used in the supplied form are computed and checked against its cached labels. Unsupported formulas are distinguished from Excel errors and retain their source caches during export.

The redistributable synthetic fixture additionally tests protection, date/currency rendering, font edits and preservation of untouched rich runs/links. Full theme/indexed-color rendering, mixed-run editing, advanced objects, general named formulas, complete validation, full-grid performance and Windows Excel screen/print comparison remain open. See [compatibility limits](COMPATIBILITY.md) and [test execution](TESTING.md). Reopen originals imported by older builds because their saved models lack source-part mappings.
