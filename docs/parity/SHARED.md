# Shared capability contracts and gap register

Every family below is an open umbrella acceptance gate. Existing subsets do not complete the family. Split it into concrete contracts using [CONTRACT_TEMPLATE.md](CONTRACT_TEMPLATE.md), and assign every relevant command from [COMMAND_COVERAGE.md](COMMAND_COVERAGE.md). Non-ribbon behavior listed here must also receive contracts. These are implementation requirements, not a claim that detailed behavior has already been researched.

## S-BASE: Reference environment and inventory

Dependencies: none. Current entry points / proposed ownership: All installed builds and catalog placements; non-ribbon workflows; explicit modern-feature supplement.

Acceptance scope: Pin executable hashes, locale, fonts, paper/printer configuration and Office options per oracle; reconcile every command and hidden/contextual workflow against live Office.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## S-MODEL: Document models, transactions and migrations

Dependencies: S-BASE. Current entry points / proposed ownership: src/model.ts; editor commands; source identities.

Acceptance scope: Version schemas; preserve unknown package nodes; make multi-object commands atomic; migrate older snapshots without losing edits; prove undo/redo and checkpoint consistency.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## S-STORAGE: Persistence, recovery and large assets

Dependencies: S-MODEL. Current entry points / proposed ownership: src/storage.ts; IndexedDB; proposed worker/OPFS assets.

Acceptance scope: Crash/restart, quota denial, eviction, migration, tab conflict and corrupted-backup tests; keep original hashes; save status follows committed transactions; recover unsaved changes or expose the loss.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## S-TEXT: Fonts, shaping, international text and typography

Current Word subset: measured uniform/empty-line advances and single-line mixed-size leading with large-mark exclusion/empty recovery and uniform single-line/wrapped excess-leading placement, with native fixtures and actual export comparisons. Exact baselines, wrapped mixed text and the cross-application shaping matrix remain open. See [W-TEXT-UNIFORM-LINES](contracts/W-TEXT-UNIFORM-LINES.md#uniform-wrapped-line-subset); its bounded milestones do not close this family.

Dependencies: S-MODEL. Current entry points / proposed ownership: Font controls; text metrics; proposed shared shaping service.

Acceptance scope: Test exact-font and documented fallback runs, complex scripts, bidi, CJK, IME, ligatures, vertical text, line breaking and font embedding restrictions across all applications.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## S-DRAW: Shared drawing and chart primitives

Dependencies: S-TEXT. Current entry points / proposed ownership: DrawingML adapters; drawing-colors.ts.

Acceptance scope: Represent complete shapes, paths, groups, transforms, paint/effects, text, chart data, themes and anchors; compare numeric geometry and rendered output with native Office.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## S-PACKAGE: OOXML, binary and alternate format coverage

Dependencies: S-MODEL. Current entry points / proposed ownership: src/formats.ts; retained package writers.

Acceptance scope: Cover DOCX/XLSX/PPTX Strict/Transitional, templates, macro-enabled packages, DOC/XLS/PPT, XLSB, ODF, CSV encodings and supported PDF workflows. Model editing separately from opaque preservation; never certify conversion solely from unchanged bytes.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## S-CLIPBOARD: Clipboard, selection and command dispatch

Dependencies: S-MODEL. Current entry points / proposed ownership: Shared command routing; editor selections.

Acceptance scope: Office-to-browser and browser-to-Office rich clipboard; Paste Special, shortcuts, contextual menus, drag/drop, touch, IME, focus and selection restoration. Permission-denied and format-fallback paths remain usable.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## S-ACCESS: Accessibility, proofing and responsive UI

Dependencies: S-TEXT,S-CLIPBOARD. Current entry points / proposed ownership: Svelte ribbons; accessibility and local dictionaries.

Acceptance scope: Keyboard-only workflows, screen-reader announcements, high contrast, zoom and all tab/dialog options. Offline spelling/grammar dictionaries, language changes and custom dictionaries require outcome tests.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## S-PRINT: Print and deterministic export

Dependencies: S-TEXT,S-DRAW. Current entry points / proposed ownership: PDF/render adapters; native PDF oracle.

Acceptance scope: Page geometry, printer-independent PDF generation, scaling, margins, page ranges, notes/handouts, headers/footers and clipping compared under pinned fonts and paper settings.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## S-AUTOMATION: Macros, extensions and browser equivalents

Dependencies: S-MODEL,S-STORAGE. Current entry points / proposed ownership: Proposed worker interpreter and capability API.

Acceptance scope: Research VBA language/object model and file-format semantics; implement bounded document automation with cancellation and permissioned file capabilities. Record COM/ActiveX/native-library gaps; a new JS API alone does not complete VBA compatibility.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## S-NETWORK: Connections, collaboration and sharing equivalents

Dependencies: S-STORAGE,S-AUTOMATION. Current entry points / proposed ownership: Proposed local adapters and optional explicit peer transport.

Acceptance scope: Local file/snapshot connections, deterministic refresh, collaborative merges and offline conflicts; evaluate peer transfer separately from hosted identities/presence. No hidden upload, service dependency or automatic email sending.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## S-PERF: Performance, isolation and release verification

Dependencies: S-STORAGE,S-PRINT,S-ACCESS. Current entry points / proposed ownership: Workers; virtualized editors; test runners.

Acceptance scope: Benchmark import/edit/recalc/render/export with recorded hardware, browser, memory and duration. Test hostile archives/XML, cancellation, unsupported features, offline use and recovery. All critical correctness gates must pass.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.
