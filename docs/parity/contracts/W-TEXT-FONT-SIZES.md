# W-TEXT-FONT-SIZES: imported run sizes and selected-text size controls

## Identity and scope

Word / W-STYLE; status **partial**. This is a bounded child scope of the font dependency in W-MODEL-SECTIONS, with no whole-command ownership and no dependency on completion of that entire parent. Related catalog controls: FontSize, FontSizeIncrease and FontSizeDecrease. The baseline is the installed x64 Word 16.0.4627.1000 executable in baseline.json.

User outcome: import run sizes within Word's 1–1638pt UI range without the former 8–160pt clamp; enter valid half-point sizes for selected text; grow/shrink selected text using the native size steps. [Microsoft's size range](https://support.microsoft.com/en-au/office/change-the-font-size-931e064e-f99f-4ba4-a1bf-8047a35552be) and [Font.Size](https://learn.microsoft.com/en-us/office/vba/api/word.font.size) were researched September 10, 2026 and checked against the installed application.

## Behavioral specification

1. Import the authored nine-paragraph fixture with sizes 1, 1.5, 6, 7.5, 10, 160, 160.5, 200 and 1638pt. All computed run sizes must match within 0.01pt CSS serialization tolerance.
2. Select the second paragraph's single character using document navigation. Enter 6, 200, 1.5 and 1638pt in sequence. Each edit must preserve other paragraphs, support Undo/Redo, survive saved reload and reimport, and export the exact half-point value accepted by native Word.
3. Enter 0, 1638.5, 1.25 or an empty value: reject the change and restore the displayed current value. This is browser validation, not a claim of identical native error-dialog wording or decimal-comma entry behavior.
4. Grow/shrink a 1.5pt selection to 2.5/1pt, with undo. Unit cases cover 21 independently measured native Font.Grow/Shrink inputs, including fractional sizes and both bounds. Below 8pt the native path advances by one point, the middle range uses its preset steps, and above 72pt it uses tens bounded by 1638pt. Growing 200pt must produce 210pt, never shrink it to 160pt.
5. The source's automatic `_GoBack` bookmark is retained. Editing its bookmarked first paragraph still hits the existing unsupported-formatting guard; the successful edited paragraph is deliberately unbookmarked. General bookmark-safe formatting remains open, not silently stripped from the fixture.

## Model and implementation

`docx-typography.ts` expands the bounded run-size conversion range; `WordEditor.svelte` validates the numeric input's range/half-point step and restores invalid entries; `word-commands.ts` implements measured Grow/Shrink steps. Existing Tiptap marks, history and retained/new DOCX writers remain the semantic representation. Source package parts are retained. No schema migration is introduced: older saved explicit 8/160pt values cannot safely be distinguished from earlier clamping without edit provenance, so this change does not overwrite them from the original file.

## Acceptance matrix

| Dimension | Cases and evidence | Status |
| --- | --- | --- |
| import/render | Nine native-saved boundary sizes; unit decorator and computed Chromium fonts | Passing subset |
| edit | Four direct size edits and two grow/shrink actions; 21 native step pairs | Passing subset |
| persist | Undo/Redo, saved reload, invalid-input nonmutation | Passing subset |
| export/native | Four actual downloads reopened in installed Word; exact run sizes and untouched paragraph sizes | Passing subset |
| create | New-document writer exists but full boundary authoring comparison is not covered here | Pending |
| interaction/errors | Selected-text numeric entry and rejection tested; native UI/dialogs, mixed selections, caret-only typing and locale matrix | Partial |
| performance | Nine-paragraph boundary fixture only; large-document performance is not certified | Pending |

Cross-browser, native dialog/keyboard variants, full font effects and damaged/out-of-range OOXML behavior remain unverified. Native checks use owned hidden COM instances, disable macros/external refresh, restore options and preserve user files.

## Evidence and reproduction

`tests/fixtures/word-font-sizes.docx` is Noffice-authored using installed Word; only personal core metadata is replaced without changing namespace prefixes. Regenerate with `scripts/verify-word-font-sizes.ps1`, then `python scripts/prepare-word-font-sizes-fixture.py`. A prior ElementTree metadata rewrite broke namespace-valued attributes; that producer failure was corrected before accepting the fixture.

After building, run:

```powershell
npx.cmd vitest run src/word-font-size.test.ts src/word-commands.test.ts src/docx-preserve.test.ts
npx.cmd playwright test tests/word-font-sizes.spec.ts --workers=1
powershell.exe -NoProfile -File scripts/verify-word-font-sizes.ps1 -Compare
```

`.local/word-font-sizes/browser-report.json` binds the tested source and four downloaded files. `native-report.json` records their exact native sizes, 21 native step pairs, executable and script hashes. Regenerate native evidence whenever browser exports or the harness change. Full command/family verification remains open.
