# Word document preservation

Imported DOCX files now retain their original package when supported paragraphs are edited. This replaces the previous reconstruction path, which discarded headers, notes, review metadata, drawings and other unmodeled content. An unchanged file still exports its exact original bytes.

## Implemented behavior

- Whole-document page choices now record explicit override intent, so applying A4/Normal to a custom imported page works even when the legacy controls already approximated it as A4/Normal. Page changes share editor Undo/Redo and persist separately from immutable section source data. A typed resolver provides eligible single-section screen/print rectangles and independent inherited header/footer references; mixed-section layout and full header/footer rendering remain open. [Current section contract](parity/contracts/W-MODEL-SECTIONS.md).

- Imported main-part paragraphs now retain a versioned source-section map, including direct header/footer references, floating-image anchors and note references. Legacy backups/snapshots recover it from retained bytes without replacing edits or invalidating unchanged-original export identity. Tracked historical sectPr settings are excluded from live page-layout import/export. This is source metadata, not an effective section editor or pagination engine. [Contract, native evidence and remaining limits](parity/contracts/W-MODEL-SECTIONS.md).

- Body paragraphs, table-cell paragraphs, footnotes and endnotes retain source identities through editor HTML and native backups. Import annotates a disposable reading copy; the original bytes are untouched.
- Direct and inherited run fonts/sizes/emphasis are resolved from document defaults, paragraph and character `basedOn` chains, style toggle properties, and Latin theme font references. Paragraph alignment, supported spacing and direction are displayed from resolved properties. Cyclic style chains are bounded.
- Plain paragraph text and supported selected formatting edits preserve existing run properties that the editor does not model, such as language and source color. Supported changes include font family/size, bold/italic/underline/strike, script marks, color and highlight. This is a subset of Word formatting.
- Paragraph indentation and pagination flags now have direct/inherited import, selected editing, individual undo steps and new/retained DOCX export. Browser printing uses CSS constraints; see [WORD_PARAGRAPHS.md](WORD_PARAGRAPHS.md) for native pagination and inherited hanging-style limits.
- Selected paragraph alignment, before/after spacing, line spacing and text direction patch paragraph properties. Whole-document paper, margin and orientation choices patch section properties; unrelated section properties remain.
- Text-only edits in supported bookmark/hyperlink paragraphs retain their surrounding XML. Simple mapped table-cell and footnote text edits update their source paragraphs.
- Simple hyperlink paragraphs additionally support selected formatting, insertion, destination edits and removal. Relationship IDs are local to each source part; shared destinations are never retargeted in place. Existing attributes, mixed run properties and leading footnote/endnote markers remain. New paragraphs and new-document exports support rich links. [Hyperlink evidence and limits](WORD_LINKS.md).
- Existing simple paragraphs support insertion, editing and removal of soft line breaks, explicit page breaks and tabs. `word-tab.ts` represents tabs as inline editor nodes so HTML whitespace normalization cannot turn them into spaces. Tab inserts a tab outside tables; table navigation retains its existing behavior. Source run formatting and unrelated package parts are retained. Browser tab display uses approximate four-column stops; custom Word tab-stop layout remains open. Clearing/column breaks and edits across fields, bookmarks or other complex structures still fail explicitly. Break edits inside hyperlinks need additional acceptance coverage.
- New ordinary paragraphs, blank paragraphs, rich text and inline breaks are inserted at their document position. Final section properties remain last; section breaks are not copied into new paragraphs. New list editing is restricted to mapped numbering contexts and needs broader coverage.
- Ctrl+Home/End and Ctrl+Shift+Home/End use explicit editor selections.
- Export changes only the affected source XML parts. All other ZIP entry payloads remain intact. It never silently falls back to rebuilding an imported DOCX.
- Retained bytes restored from native backups pass archive expansion checks before source parsing, just like ordinary Office imports.

## Validation

`src/docx-preserve.test.ts` covers retained payloads, text/run edits, bookmarks, new rich/blank paragraphs, section boundaries, mapped table cells and rejection of unsupported field/table edits. `src/docx-styles.test.ts` covers defaults, style ancestry, direct overrides, theme fonts, toggle semantics and cyclic chains.

`tests/docx-fidelity.spec.ts` exercises the actual static application: selected underline, text editing, spacing/direction controls, footnote editing, document navigation, soft breaks/tabs, export and reimport. It compares every unrelated ZIP payload and asserts source page settings remain. The 20 external DOCX scenarios include 19 accepted documents and one deliberate excessive-nesting rejection; all accepted paragraph-addition outputs have no counted package, text or feature decreases. See [the corpus evidence](OFFICE_CORPUS.md).

## Remaining limits

This is not complete desktop Word functionality or native layout certification. Exact pagination, headers/footers on screen, tracked-change authoring/acceptance, field editing, rich object editing, content controls, equations, citations, complete list/table restructuring, style management, complex-script theme selection and font embedding remain open. Imported heading-style changes and many structural edits are rejected at export. Some paragraphs produced by drawing/review conversion are opaque: they can remain unchanged, but cannot yet be edited safely. Paragraph splitting across complex source boundaries needs more coverage.

The font resolver does not implement the complete table-style/conditional-formatting cascade, all script-specific fonts, or every style-property interaction. New paragraphs do not reproduce every inherited Word paragraph/run property. Large-paragraph edits are capped at 200,000 characters; import XML and archive limits still apply.

Reopen documents imported by older builds to obtain paragraph mappings. Keep a native backup before reopening. The authored hyperlink browser export now passes native Word destination/formatting checks and PDF rendering; broader native corpus/layout validation remains incomplete. The full desktop parity inventory remains unchecked in [TASKS.md](../TASKS.md).

The soft-break/tab browser workflow uses actual Shift+Enter and Tab keystrokes, saves/reloads, inspects exported `w:br`/`w:tab` elements and reimports the download. Unit tests cover source run properties, unaffected parts, break insertion/removal and rejection of clearing-break conversion. Reopen older imports to obtain explicit tab nodes.

## Explicit page boundaries

`word-page-break.ts` adds an inline page-break node, Ctrl+Enter and Insert > Page break. Import marks the exact source break position in a disposable reading copy because the semantic converter otherwise omits it. The retained writer keeps the break within its original paragraph and emits `w:br w:type="page"`; the new-document writer also emits a real page break. Surrounding text/style edits and page-break insertion/removal have preservation tests.

The editor displays a labeled boundary in continuous layout. Browser printing hides that label and applies a page break. The Chromium workflow verifies insertion by keyboard and toolbar, undo, save/reload, unchanged unrelated ZIP payloads, exported/reimported boundaries and a two-page PDF. Local artifacts are `.local/word-validation/page-break.docx` and `page-break.pdf`. This is selected browser pagination evidence, not native Word pagination parity. Headers, footnote placement, complex paragraph structures and exact line/page metrics still need a full layout engine. Reopen older imports to load their explicit page-break nodes.

Primary reference: [Microsoft WordprocessingML break element](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.break?view=openxml-3.0.1).
