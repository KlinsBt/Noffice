# Noffice task checklist

**Desktop Word/Excel/PowerPoint parity: incomplete.** Checked items certify only the stated tested scope. The authoritative inventory remains **5,579 command identities, 11,206 placements and 51 capability families**, with no certified parity percentage.

## Latest update - September 16, 2026

**Active work:** the existing section insertion/scoped-command/container/tracked/clipboard requirement. Source-derived fresh DOCX numbering and native round trips now pass the stated flat-list profile. Continue table row insertion, source identities, AutoFit widths/padding and both writers/layout; tracked edits and rich clipboard follow. The original parent stays open.

**Implemented in the active pagination requirement:** header/footer import, editing, page options, linking, section-deletion repair and empty-slot creation, including private missing styles, kerning inheritance and wrapped baseline repairs. Evidence applies to the stated cases; broader requirements remain open. Detailed implementation and historical checkpoints are in the linked contract.

**Updated September 16:** Fresh DOCX copies now retain qualified list marker fonts, sizes, indents and starts across section fragments, with independent counters and guarded failure recovery. Verification passes **4,226 unit tests / one existing skip**, zero diagnostics, a 54-file static build and **139 distinct browser workflows**. All **200 native DOCX comparisons / 400 page pairs** and **120 direct PDFs / 240 pages** pass their unchanged gates, with reviewed complete-page evidence. All 40 CSS-print files retain documented horizontal failures. Tables, tracked edits, rich clipboard and the complete acceptance matrix remain open. [Evidence and exact handoff](docs/parity/contracts/W-MODEL-SECTIONS.md#fresh-list-continuation-september-16).

**Still open:** full story/section authoring and references, the broader page-option and link/clone acceptance, broader fonts/scripts/formatting/selection, storage and conversion failure combinations, cross-browser/print acceptance and the complete original pagination matrix. Thirty Inter horizontal PDF cases and forced-column horizontal error0.256pt remain; CSS print retains the earlier A4 width0.36pt and0.154pt/0.248pt horizontal differences. The inventory remains193 production packages with eight unresolved full notices and further bundled-component review. No parity percentage, parent completion or command certification is claimed.

[Exact active handoff and artifact paths](docs/parity/contracts/W-MODEL-SECTIONS.md#fresh-list-continuation-september-16) · [Historical delivery evidence](TASK_HISTORY.md).

## How to use this checklist

This is the active work list, not a session log. Update existing rows in place. Put test details and exact handoffs in the relevant feature contract; retain earlier delivery notes in [TASK_HISTORY.md](TASK_HISTORY.md). Consolidating repeated historical rows does not complete or remove any product requirement. All original granular requirements remain there and in the [family register](docs/parity/README.md) and [command inventories](docs/ribbon/README.md).

Check a task only after implementation, relevant tests, independent Office comparisons where required, and `npm.cmd run parity:check` pass. Follow [LLM_WORKFLOW.md](docs/parity/LLM_WORKFLOW.md) and [ENGINE_PLAN.md](docs/parity/ENGINE_PLAN.md). Passing build/test infrastructure is separate from product completeness.

For the active one-checkbox workflow, record bounded progress and evidence within its existing row and contract. Do not create substitute subset checkboxes. Check that original requirement only when its complete acceptance scope passes. Checkbox counts mix historical milestones and broad requirements and must not be used as a completion percentage.

## Verified implementation milestones

- [x] Excel imported relative-name subset: bind the measured A1-origin convention at each formula caller, including mixed absolute axes, workbook/local aliases, tested range consumers and cross-cell name chains. Precedent/caller edits, division-error recovery, Undo/Redo, reload, raw export caches and three independently compared native XLSX exports pass. The tested scalar intersection subset now has native evidence; complete arrays, relative-sheet/whole-axis names, static name binding and workers remain open. [Evidence and limits](docs/parity/contracts/X-MODEL-DEPENDENCIES.md#relative-name-anchor-subset-active-september-11-2026).

- [x] PowerPoint rotated/reflected child-frame subset: move/resize the tested unrotated leaves through outer rotation/reflection and normalize measured rotated-child quadrant bounds. Numeric Width/Height changes now retain the native rotated corner. History/reload/slideshow, invalid/unsupported recovery, legacy migration, actual exports/reimports and native re-edit/frame/PDF comparisons pass. Broader child transforms, coarse rotation/reflection, grouped text, whole-group authoring and the full native interaction matrix remain open. [Evidence and limits](docs/parity/contracts/P-MODEL-INHERITANCE.md#quadrant-bounds-and-numeric-size-continuation).


- [x] Excel literal static-reference subset: preserve absolute/mixed endpoints and sheet IDs; index cell/range/full-row/full-column references compactly for retained cell-edit and table-rename export recalculation. Inactive branches remain unevaluated; incomplete bindings and work-budget overflow stay conservative. Native edit/delete/branch exports, raw caches and regression checks pass. [Evidence and limits](docs/parity/contracts/X-MODEL-DEPENDENCIES.md#static-reference-and-blank-cell-subset).
- [x] Excel scalar blank-reference subset: deleting a referenced cell produces native numeric zero while referenced empty-text formulas remain strings. Tested blank/type/statistical consumers, division-error recovery, Undo/Redo, reload and three actual exported XLSX comparisons pass. Complete function/type/array semantics remain open. [Evidence and limits](docs/parity/contracts/X-MODEL-DEPENDENCIES.md#static-reference-and-blank-cell-subset).


- [x] Word uniform wrapped-line subset: preserve the full source text width and hanging wrap-spaces; match seven soft-wrap and three hard-break identities across source/Double/minimum/Single. Controlled first-line leading, history, invalid input, reload/export/reimport, 50% zoom and pointer typing pass; all three actual exports match native properties and exact PDFs. Nine unwrapped empty hard-break/paragraph-mark cases now add native advances, terminal typing, pasted-font history, real 50% caret/fallback recovery and four independently compared DOCX/PDF exports. An independent paragraph-background probe now verifies 40 per-state advances and six relative-leading shifts within the original 0.15pt bound, preserving glyph matrices exactly. The original four mixed-origin diagnostics and absolute glyph parity remain open. [Evidence and limits](docs/parity/contracts/W-TEXT-UNIFORM-LINES.md#paragraph-coordinate-and-mixed-wrap-continuation).

- [x] Word uniform single-line leading placement: tested 10pt text stays fixed when switching automatic multiples; a 50pt minimum puts excess above it. Double/minimum/Single edits, invalid input, Undo/Redo, reload, wrapping recovery and three actual exports match independent native comparisons. Section measurements also refresh after deferred font changes, preserving long-flow fallback and Undo recovery. Exact baselines and wrapped-line placement remain open. [Evidence and limits](docs/parity/contracts/W-TEXT-UNIFORM-LINES.md#uniform-single-line-leading-placement).

- [x] Word large paragraph-mark subset: three 40pt-mark cases exclude the mark from eligible nonempty single-line advances; clearing and typing restore 69pt empty/40pt-text advances. Size editing, Undo/Redo, reload, actual export/reimport and all three independent native property/PDF comparisons pass. Wrapped text and exact baselines remain open. [Evidence and limits](docs/parity/contracts/W-TEXT-UNIFORM-LINES.md#large-paragraph-mark-subset).

- [x] Word bounded mixed-size line-advance subset: the existing Arial single-line cases and same-font Latin soft wrapping match native paragraph/line-box geometry. The regular Arial mixed-wrap subset now also matches 66 paragraph-relative absolute/relative DOM glyph baselines and 66 baselines in six actual Chromium PDFs at the unchanged 0.15pt bound. Spacing/font edits, typing, invalid input, Undo/Redo, reload, fallback recovery, 50% pointer typing and five actual DOCX exports/reimports plus native re-edit pass. Other fonts/scripts, mixed hard breaks, glyph outlines, whole-page coordinates and full save-session semantics remain open. [Evidence and limits](docs/parity/contracts/W-TEXT-UNIFORM-LINES.md#bounded-native-glyph-baseline-continuation).

- [x] Excel imported scoped-name binding subset: workbook aliases retain workbook scope despite local shadows; quoted-sheet names and internal `[0]!Name` references calculate correctly. Precedent edit/delete, Undo/Redo, reload and both actual XLSX exports match native values, types, formulas and name definitions. Scope-aware cycle and unsupported-reference tests pass. [Evidence and limits](docs/parity/contracts/X-MODEL-DEPENDENCIES.md#scoped-name-binding-subset).
- [x] Excel legacy name/array metadata migration: restore missing source scopes, A1 origins and native array identities while preserving distinct precedent edits, original bytes and storage revision. Unit and seeded IndexedDB browser tests pass; mismatched names and ambiguous changed array anchors/members reject with backup/reopen guidance. Full name and multi-cell array authoring remain open. [Evidence and limits](docs/parity/contracts/X-MODEL-DEPENDENCIES.md#single-cell-cse-identity-and-conditional-array-continuation).

- [x] PowerPoint nested-group frame subset: compose source transforms for 12 native rectangle/ellipse frames across nested scale/translation, outer rotation and horizontal reflection. Model persistence, canvas and slideshow checks pass; all four actual exported slides match native properties and rendered PDFs. Skew and grouped text metrics remain open. [Evidence and limits](docs/parity/contracts/P-MODEL-INHERITANCE.md#nested-group-transform-subset).
- [x] PowerPoint axis-aligned group-child edit subset: tested move/resize edits invert source coordinates and normalize affected group containers. Eligible coarse children now snap to the native-measured retained grid; field commits, pointer/keyboard, no-op/invalid-input recovery, history, reload, actual export/reimport and native re-edit/rendering pass. Unrelated payloads survive. Coarse rotation/reflection, child transforms, grouped text, skew and full interaction parity remain open. [Evidence and limits](docs/parity/contracts/P-MODEL-INHERITANCE.md#coarse-coordinate-authoring-continuation).
- [x] PowerPoint legacy group-coordinate migration: recover effective frames and missing source slide dimensions from retained transforms while preserving distinct edits, source bytes and revision identity. Unit and seeded IndexedDB browser tests pass, including unchanged byte-identical export and edited coarse-grid recovery. [Evidence and ambiguity boundary](docs/parity/contracts/P-MODEL-INHERITANCE.md#coarse-coordinate-authoring-continuation).

- [x] PowerPoint inherited-placeholder subset: two distinct native master/layout chains resolve the measured title fonts, alignment and geometry; source placeholder and group ancestry IDs survive model persistence. Title text/position edits, Undo/Redo, reload and actual PPTX/native rendering comparisons pass. [Evidence and limits](docs/parity/contracts/P-MODEL-INHERITANCE.md).
- [x] PowerPoint retained rich-run subset: canvas/previews/slideshow display the tested mixed size, font, bold, italic, underline and color. Replacing a whole run retains that run's style; history, reload, actual export/reimport and independent native character/PDF comparisons pass. [Evidence and limits](docs/parity/contracts/P-MODEL-INHERITANCE.md).
- [x] PowerPoint legacy text migration: recover inherited defaults and source runs while preserving distinct saved edits, source IDs and original bytes. Unit and IndexedDB browser checks pass, including unchanged export and no migration-only storage revision. [Migration evidence and ambiguity boundary](docs/parity/contracts/P-MODEL-INHERITANCE.md).
- [x] PowerPoint first-notes package subset: create/register a missing notes master with a separate theme copy, reuse it across notes slides, and reject a broken theme target. Notes editing, Undo/Redo, reload and actual combined image/notes export/reimport pass; installed PowerPoint opens and renders the output. Full notes-page design/print parity remains open. [Native regression evidence](docs/parity/contracts/P-MODEL-INHERITANCE.md#native-regression-findings).

- [x] Cold-cache CI catalog bootstrap: download and verify the pinned catalogs before checking generated inventories. The isolated empty-cache reproduction and all three inventory checks passed; the hosted Actions job itself has not been run. [Implementation and reproduction evidence](docs/parity/contracts/W-MODEL-SECTIONS.md#current-handoff-font-bounds-uniform-leading-and-excel-invalidation).
- [x] Word font-size subset: import and edit 1-1638pt sizes, reject the tested invalid inputs, and match 21 native Grow/Shrink pairs. Editing, Undo/Redo, reload, actual DOCX export/reimport and four native export comparisons passed. [Font-size contract and evidence](docs/parity/contracts/W-TEXT-FONT-SIZES.md).
- [x] Word measured line-advance subset: match the native 10pt uniform 1.5-line advance and the empty 30pt paragraph-mark minimum advance. Targeted spacing/typing/history/reload and native export regressions passed. Mixed-run baselines and variable pagination remain open. [Uniform-line contract and evidence](docs/parity/contracts/W-TEXT-UNIFORM-LINES.md).
- [x] Excel observed dependency/reference subset: preserve caller-aligned scalar intersection, blank/error types, names, INDEX/INDIRECT and measured shaped SUMPRODUCT arithmetic. The bounded single-cell CSE/conditional-array subset now preserves entry identity through keyboard editing, Undo/Redo, reload, actual XLSX export/reimport and native re-edit, with raw saved caches and exact native PDFs. Full function/array matrices, multi-cell CSE/spills, table selectors, static names and workers remain open. [Evidence and limits](docs/parity/contracts/X-MODEL-DEPENDENCIES.md#single-cell-cse-identity-and-conditional-array-continuation).

## Current dependency queue

1. **Word: [W-MODEL-SECTIONS](docs/parity/contracts/W-MODEL-SECTIONS.md#fresh-list-continuation-september-16), partial; same active section/container requirement.** The40 source-derived fresh-list cases now pass export, native X/browser Y edits and current recovery checks. Continue the 12 measured table row cases, first resolving native AutoFit width/padding changes; then implement guarded splitting, reference repair, both writers and layout, followed by tracked edits and rich clipboard. Preserve the earlier native refresh gap and font/print/UI uncertainty.
2. **Excel: [X-MODEL-DEPENDENCIES](docs/parity/contracts/X-MODEL-DEPENDENCIES.md#cse-verification-and-exact-handoff), partial.** Preserve the 96 ordinary/CSE snapshots, single-cell keyboard/history/recovery and actual-export subset, all 39 scalar-reference and 48 shaped-reference cases. Continue multi-cell CSE identities/commands/results, complete conditional/function lifting and lazy/error/type/locale matrices, fresh-workbook native exports, spills, table selectors, static names and workers.
3. **PowerPoint: [P-MODEL-INHERITANCE](docs/parity/contracts/P-MODEL-INHERITANCE.md#quadrant-bounds-and-numeric-size-continuation), partial.** Preserve the passing quadrant bounds, native numeric-size coupling and fine/coarse subsets with unchanged precision guards. Continue coarse rotated/reflected editing, child transforms, grouped text and whole-group authoring; the full Size-pane/units/modifiers, shape-type and browser/native interaction matrices remain open.


## Word section contract progress

- [x] Source identities, legacy snapshot migration and effective geometry/inherited references; independently tested source and exported DOCX behavior.
- [x] Exact eligible single-section geometry and whole-document layout history; measured mixed-section surfaces and exact-spacing overflow, including tested widow/orphan toggles and empty continuation-line caret editing; Undo/Redo, reload, DOCX export/reimport and native page/print comparisons pass. Minimum/exact/multiple spacing authoring, omitted font fallback, paragraph-mark fonts, legacy spacing/font migration and empty-paragraph typing also have scoped browser/native comparisons.
- [x] Live section ranges, selection indicators and formatting marks; tested top-level split/join history, persistence and retained header/footer repair. Exported split/join files match Word on three rendered pages.
- [ ] Complete pagination beyond the tested exact-spacing path: automatic/at-least line-box and baseline metrics, broader widow/keep/font/selection combinations, continuous/column breaks, headers/footers and the full print matrix. **Verified progress:** selected-boundary deletion, paragraph-mark/font inheritance, leading/first-line/hanging-indent continuations and precise wrapped/mixed-face baselines. 4,011 units and 152 editing/recovery/native-return browser workflows pass. Actual native file comparisons and all 80 case reviews are complete, with failed strict caret/implicit further-edit/glyph checks preserved in [Important browser uncertainties](docs/parity/BROWSER_LIMITATIONS.md). Full pagination, native keyboard, broader fonts/stories/selection and print acceptance remain open. [Exact evidence and next action](docs/parity/contracts/W-TEXT-UNIFORM-LINES.md#current-executable-handoff); [history](TASK_HISTORY.md).
- [ ] Section insertion/scoped commands, container/tracked structural edits and rich section clipboard semantics. **September 16 progress:** Fresh DOCX writing now preserves qualified source list definitions, independent starts and continuation across section fragments. Native/browser typing, history, reload, actual files and export failures pass. Current verification:4,226 units/one skip, 139 browsers,200 exact native comparisons and 120 strict direct PDFs. Forty CSS-print failures remain Important. **Next:** table row splitting with native AutoFit geometry, retained/fresh writers and pagination; tracked structural edits and rich clipboard. The original requirement stays open. [Evidence and exact handoff](docs/parity/contracts/W-MODEL-SECTIONS.md#fresh-list-continuation-september-16).
- [ ] Complete native UI, cross-browser, recovery and performance acceptance for the section contract.

Word evidence, limitations and its next action live in [the current Word handoff](docs/parity/contracts/W-TEXT-UNIFORM-LINES.md#current-executable-handoff). Current shared regression totals are in the dated summary above. Historical native runs are not proof for subsequent changes.

## Known release evidence gaps

The strict corpus gate still rejects the deeply nested DOCX case and the protected-XLSX editing evidence gap. Font packaging/fallback, private price-schedule screen/print fidelity, broader native corpus checks and the complete browser/accessibility matrix remain required. Detailed earlier cases and local receipts remain in [the archive](TASK_HISTORY.md#next-tasks-to-implement-and-prove) and [TESTING.md](docs/TESTING.md).

## Remaining desktop Word scope

- [ ] Paginated layout engine matching Word line, paragraph, and page breaking.
- [ ] Full paragraph/run style inheritance and theme/font substitution.
- [ ] Sections, columns, orientation changes, headers/footers, and page numbering.
- [ ] Footnotes/endnotes, citations, bibliography, captions, indexes, and tables of contents.
- [ ] Cross-references, fields, equations, mail merge, and content controls.
- [ ] Track changes, comments, compare/combine, and revision metadata preservation.
- [ ] Advanced tables, nested tables, row splitting, repeat headers, and text wrapping.
- [ ] Floating images, anchors, drawings, shapes, SmartArt, charts, and embedded objects.
- [ ] Full proofing, accessibility checking, language tools, and document protection.
- [ ] Word-authored DOCX corpus with semantic and visual round-trip checks.

## Remaining desktop Excel scope

- [ ] Complete function inventory, type coercion, errors, operator semantics, and locale behavior.
- [ ] Dynamic arrays, named ranges, structured references, tables, and formula auditing.
- [ ] Date systems, financial/statistical/engineering functions, and all lookup functions.
- [ ] Complete dependency graph scheduling, worker calculation, cancellation, and large-model performance beyond the checked observed-dependency subset.
- [ ] Row/column insert/delete/move with reference repair; merges, resize, freeze, hide, and grouping.
- [ ] Conditional formatting, validation, comments, hyperlinks, and protected ranges.
- [ ] Charts and chart formatting, pivot tables, pivot charts, slicers, and timelines.
- [ ] Goal Seek, Solver, scenarios, data tables, Power Query, and Power Pivot equivalents.
- [ ] External-data connections and browser-compatible alternatives to desktop integrations.
- [ ] Print areas, repeating rows, page setup, and workbook/worksheet protection.
- [ ] Excel-authored XLSX corpus with calculation and visual-fidelity checks.

## Remaining desktop PowerPoint scope

- [ ] Master slides, layouts, placeholders, themes, inheritance, and arbitrary aspect ratios.
- [ ] Rich text runs, fonts, bullets, autofit, typography, and exact text metrics.
- [ ] All shapes, connectors, groups, rotation, cropping, guides, and alignment/distribution.
- [ ] Tables, charts, SmartArt, equations, embedded objects, and editable vector graphics.
- [ ] Transition and animation timelines with trigger and timing semantics.
- [ ] Audio/video, narration, recording, captions, and media trimming.
- [ ] Presenter view, multi-display behavior, custom shows, sections, and rehearsal timings.
- [ ] Comments, review, protection, and metadata-preserving round trips.
- [ ] PDF/video/image export and handout/notes printing.
- [ ] PowerPoint-authored PPTX corpus with visual/semantic/presentation validation.

## Remaining cross-suite release gates

- [ ] Package-preserving OOXML editing that retains unsupported relationships and parts.
- [ ] Legacy DOC/XLS/PPT, macro-enabled files, encrypted documents, and ODF strategy.
- [ ] Browser-compatible macro/add-in automation and a documented boundary for Windows-only APIs.
- [ ] Peer collaboration and conflict merging without a mandatory application backend.
- [ ] Worker isolation, cancellation, streaming/bounded decompression, and adversarial import corpus.
- [ ] Full Firefox/WebKit matrix, screen-reader audit, IME/RTL testing, and touch editing review.
- [ ] Quota exhaustion, storage eviction, crash recovery, migration, and multi-tab upgrade testing.
- [ ] Complete transitive license/provenance review; resolve packages with missing or unknown bundled notices before public distribution. The pinned saxes notice is now shipped; seven of 187 packages remain unresolved. The new offline check rejects missing/stale notices; three verifier tests pass. [Evidence and remaining packages](docs/licenses/README.md).
- [ ] Independent Windows Office open/save/render checks and feature inventory signed off against a specific Office version/build.
- [ ] Full desktop functionality parity for all three apps demonstrated by evidence.
