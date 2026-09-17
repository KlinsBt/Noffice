# Search and replacement

The three editors now share a local search panel with Find next/previous, Replace, Replace all, Match case, result counts and Escape to close. Ctrl+F and Ctrl+H open search from the editor. The panel uses the responsive ribbon's available-width approach and keeps controls reachable in narrow windows.

## Implemented behavior

Word searches contiguous text across bold/italic/font/link run boundaries within paragraphs, headings and table cells. Whole-word matching uses Unicode letters, numbers and combining marks. Replacement inherits the first matched character's marks, retains paragraph/source identities, and creates one undo step. Searches stop at paragraph boundaries and inline objects or breaks. It does not search headers/footers or implement Word's wildcard, formatting, special-character or language-specific search options.

Excel searches the current sheet, visible sheets in a workbook, or the selection captured when the panel opens. It supports row/column order, case, entire-cell matching, formulas/raw contents, displayed values and notes. It supports `*`, `?` and `~` escapes, with a literal-mode checkbox. Find next/previous continues relative to the active cell and wraps. Hidden sheets, rows and columns are excluded explicitly. Replacement operates on formulas/raw contents; it never writes into a calculated result. All target cells are validated before committing: a locked cell, failed validation or oversized value cancels the entire replacement. Cell style, source metadata and unrelated cells remain intact; changed values invalidate their old types/caches.

PowerPoint searches occurrences in editable text-bearing objects across slides, navigates to the object and selects the occurrence in its text inspector where available. Replacement preserves object IDs, geometry, direct formatting and speaker notes, with a single undo step. Masters, charts, tables, embedded objects and notes are not searched. Existing imported-export guards still apply to unsupported field/soft-break or paragraph-structure edits.

Excel's Home alignment group additionally exposes Wrap text and vertical top/middle/bottom alignment. Changes apply to the selected range and survive XLSX export and reload. Automatic row-height fitting remains unfinished.

## Bounds and interoperability

Search accepts up to 256 query characters, 8 million scanned characters and 10,000 matches per operation. Wildcards use dynamic programming with a four-million-step budget rather than backtracking regular expressions. Empty wildcard matches are omitted. Replacement uses literal strings, including dollar signs, and rejects results exceeding 2 million document text characters, 32,767 characters per spreadsheet cell or 100,000 characters per slide object before committing. Locale-specific collation is not certified against every Office language/version.

Unit evidence: `text-search.test.ts`, `search-wildcards.test.ts`, `word-search.test.ts`, `office-search.test.ts`. Browser evidence: `tests/search.spec.ts` checks formatted DOCX matches, selection, single/all replacement, undo/redo, workbook scope, formula/value separation, protected-cell cancellation, wildcard escapes, slide navigation, saved reload and real DOCX/XLSX/PPTX downloads. DOCX and PPTX tests compare untouched ZIP payloads. XLSX checks source styles, changed formulas, both worksheets and alignment metadata. This evidence does not certify full desktop search or overall Office parity.

## Native Excel comparison

On September 9, 2026, installed Microsoft Excel 16.0 build 4627 evaluated 23 authored cases using both `Range.Find` and `Range.Replace`. The receipt in `tests/fixtures/native-excel-search.json` binds the input bytes with SHA-256 `fbc5e2180808b2266af3c6b3472c1972059701378ffb70c5bf371a197ad8803c`. `src/native-search.test.ts` checks that binding and compares every native find result and replacement string.

These experiments corrected two initial assumptions: an internal `*` takes the shortest successful span (`a*b` replaces both spans in `a1b a2b`), while a trailing `*` consumes the remaining text; `~` escapes ordinary characters as well as wildcard characters. Native Find also locates a dangling `~` literally, while native Replace leaves that query unchanged. The matcher distinguishes finding from replacement to retain this observed behavior. These are measured results for the recorded cases and Office build, not a claim about all wildcard, Unicode or locale combinations.

To regenerate the evidence with Windows Excel installed:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-excel-search.ps1
npx.cmd vitest run src/native-search.test.ts src/search-wildcards.test.ts
```

The script creates and closes a temporary workbook in a separate hidden Excel instance. It does not open user documents. Native automation and fixtures are development tools and are excluded from the client-side application.

## Primary references

- [Microsoft: Excel Find and Replace](https://support.microsoft.com/office/find-or-replace-text-and-numbers-on-a-worksheet-0e304ca5-ecef-4808-b90f-fdb42f892e90) documents scope, traversal order, match options, wildcards, and the distinction between searching values and replacing formulas.
- [Microsoft: Find text in Word](https://support.microsoft.com/en-us/word/find-text-in-a-document) documents case and whole-word filters and Ctrl+F.
- [Microsoft: PowerPoint Find and Replace](https://support.microsoft.com/en-us/powerpoint/find-and-replace-text) documents single-occurrence and all-occurrence replacement.
- [Microsoft: PowerPoint keyboard shortcuts](https://support.microsoft.com/en-us/accessibility/powerpoint/use-keyboard-shortcuts-to-create-powerpoint-presentations) documents Ctrl+F/Ctrl+H. Platform-specific shortcut variants remain to be validated.

Research checked September 9, 2026. The reference repository's IndexedDB design was reviewed as architectural context; its MPL source was not copied. Replacements use Noffice's existing revision-checked transaction and original-package preservation paths.
