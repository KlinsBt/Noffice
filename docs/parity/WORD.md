# Word capability contracts and gap register

Every family below is an open umbrella acceptance gate. Existing subsets do not complete the family. Split it into concrete contracts using [CONTRACT_TEMPLATE.md](CONTRACT_TEMPLATE.md), and assign every relevant command from [COMMAND_COVERAGE.md](COMMAND_COVERAGE.md). Non-ribbon behavior listed here must also receive contracts. These are implementation requirements, not a claim that detailed behavior has already been researched.

The active continuation is the full existing Word pagination checkbox, including typography, selection, sections, headers/footers and print. Page-option/restoration/recovery subsets have reviewed actual-file evidence; header/footer link/clone authoring is undergoing native end-to-end comparison. Keep all family gates open until their complete matrices pass. [Current scope, failures and exact next action](contracts/W-TEXT-UNIFORM-LINES.md#current-executable-handoff).

## W-STYLE: Styles, font/paragraph dialogs and themes

Dependencies: S-TEXT. Current entry points / proposed ownership: Tiptap marks; DOCX styles.xml and theme.xml.

Acceptance scope: Create/modify/delete styles, inheritance and direct overrides, keep-with-next, borders/shading and all font effects; edit then reopen in both applications without changing unrelated runs.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## W-FLOW: Pagination, sections and columns

The current line-metric prerequisite includes single-line mixed sizes, large-mark/empty recovery, uniform wrapping, hanging spaces and nine empty hard-break cases. The regular Arial mixed-wrap subset now adds 66 DOM baseline checks and 66 baselines in six actual Chromium PDFs, alongside native identities/boxes, history/reload/fallback/50% caret behavior, five actual DOCX exports and native re-edit. A separate 28-case native font/spacing profile is guarded; broader fonts/platforms/printers, uniform absolute baselines, scripts/mixed hard breaks, glyph outlines and pagination remain open. All parent gates below remain unchecked. The latest user override keeps the full Word pagination checkbox active; do not switch to another application after a bounded subset. [Scope and evidence](contracts/W-TEXT-UNIFORM-LINES.md#bounded-native-glyph-baseline-continuation).

Dependencies: W-STYLE. Current entry points / proposed ownership: Proposed layout engine over semantic document model.

Acceptance scope: Continuous/next/odd/even sections, column flow, widows/orphans, hyphenation, line numbering, page geometry and line/page breaks; compare line/page positions and print PDFs with native Word.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## W-LISTS: Lists and outline numbering

Dependencies: W-STYLE. Current entry points / proposed ownership: DOCX numbering.xml; editor list model.

Acceptance scope: Restart/continue, multilevel numbering, custom formats and style links across copy/paste and section changes; preserve numbering identities and native indentation.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## W-TABLE: Complete table editing and layout

Dependencies: W-FLOW. Current entry points / proposed ownership: Tiptap tables; DOCX grid model.

Acceptance scope: Nested tables, merged/split cells, repeated headers, row breaks, auto-fit, text direction, formulas, sorting, captions, borders and selection; fix current deep-table import gap without removing bounded parsing.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## W-OBJECT: Images, shapes, charts and embedded objects

Dependencies: W-FLOW,S-DRAW. Current entry points / proposed ownership: DOCX drawing anchors and object model.

Acceptance scope: Inline/floating wrapping, anchors, grouping, cropping, effects, SmartArt/charts and object updates; compare text reflow around drawings and retained relationship integrity.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## W-PAGES: Headers, footers and page design

Dependencies: W-FLOW. Current entry points / proposed ownership: DOCX section/header/footer relationships.

Acceptance scope: First/odd/even headers, section linking, page fields, watermarks, page borders, cover pages and backgrounds; author and edit every variant, not just retain imported parts.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## W-FIELDS: Fields, references and scholarly tools

Dependencies: W-FLOW,W-LISTS. Current entry points / proposed ownership: Proposed field engine; bookmarks and note parts.

Acceptance scope: TOC, indexes, cross-references, captions, citations/bibliography, foot/endnotes and equations; update fields after reordering, insertion and deletion; compare native results and round trips.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## W-REVIEW: Comments, tracked changes and comparison

Dependencies: W-STYLE,S-MODEL. Current entry points / proposed ownership: Revision-aware document transactions.

Acceptance scope: Author/reply/resolve comments, track insert/delete/format/move, accept/reject selected/all, compare and combine documents with stable anchors and author/time metadata.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## W-MERGE: Mail merge, envelopes and labels

Dependencies: W-FIELDS,S-STORAGE. Current entry points / proposed ownership: Proposed recipient tables and merge evaluator.

Acceptance scope: CSV/XLSX recipient import, filtering, sorting, merge fields/rules, preview and generation, envelope/label paper layouts; generate local outputs without silently invoking mail services.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## W-FORMS: Content controls, forms and protection

Dependencies: W-REVIEW. Current entry points / proposed ownership: DOCX controls/protection parts.

Acceptance scope: Text/date/list/check controls, bindings, restricted editing and document protection; validate allowed edits and preserved controls, including legacy form fields.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## W-VIEWS: Navigation, search and document views

Dependencies: W-FLOW,W-FIELDS,S-ACCESS. Current entry points / proposed ownership: WordEditor.svelte; word-search.ts.

Acceptance scope: Advanced/wildcard/format search, navigation, outline/draft/read/print views, zoom, split windows and document properties; selections and undo remain consistent between views.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## W-AUTOMATION: Word automation and output workflows

Dependencies: W-MERGE,W-FORMS,S-AUTOMATION,S-PRINT. Current entry points / proposed ownership: Word capability adapter.

Acceptance scope: Record/replay document actions, compatible object operations and bulk generation; keep unavailable OS integrations as explicit equivalent gaps; qualify print and all Word format contracts.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.
