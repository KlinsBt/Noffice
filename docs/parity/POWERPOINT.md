# PowerPoint capability contracts and gap register

Every family below is an open umbrella acceptance gate. Existing subsets do not complete the family. Split it into concrete contracts using [CONTRACT_TEMPLATE.md](CONTRACT_TEMPLATE.md), and assign every relevant command from [COMMAND_COVERAGE.md](COMMAND_COVERAGE.md). Non-ribbon behavior listed here must also receive contracts. These are implementation requirements, not a claim that detailed behavior has already been researched.

The bounded fine/coarse sibling-edit subsets now include measured rotated-child quadrant bounds and native numeric Width/Height corner retention. Native angle discovery, actual Position/Size-pane entry/history, legacy recovery, actual exports and native re-edit/frame/PDF comparisons pass. All broader families below stay open. Next PowerPoint dependencies include coarse rotated/reflected editing, child transforms, grouped text and whole-group authoring; full shape-type, units/modifiers and Size-pane behavior remain required. [Current evidence and queue](contracts/P-MODEL-INHERITANCE.md#quadrant-bounds-and-numeric-size-continuation).

## P-TEXT: Rich text, paragraphs and text fitting

Dependencies: S-TEXT. Current entry points / proposed ownership: SlideElement text model; PPTX rich runs.

Acceptance scope: Mixed runs, bullets/numbering, tabs, paragraph spacing, autofit, vertical text, bidi and hyperlinks; native text geometry and per-run typography must survive edits.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## P-MASTER: Themes, masters, layouts and placeholders

Dependencies: P-TEXT. Current entry points / proposed ownership: DrawingML inheritance and layout model.

Acceptance scope: Create/edit/apply master/layout/theme, placeholder inheritance, reset and backgrounds; preserve overrides and compare every slide plus notes/handout masters.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## P-DRAW: Shapes, images, groups and connectors

The existing child-frame subset now normalizes the tested nested containers and matches actual native outer-group selection bounds after export/reopen. This does not complete group creation/selection/authoring, rotated-child bounds, skew or grouped text. [Current bounded evidence](contracts/P-MODEL-INHERITANCE.md#group-bound-normalization-subset-september-11-2026).

Dependencies: P-MASTER,S-DRAW. Current entry points / proposed ownership: Slide scene graph; slide-arrange/stack.

Acceptance scope: All shapes/freeform paths, group-child edits, connectors, align/distribute/order, crop/effects and selection pane; compare transforms and displayed grouping after export.

The bounded [rotated/reflected child-frame edit subset](contracts/P-MODEL-INHERITANCE.md#rotated-and-reflected-child-frame-editing-subset) extends existing orthogonal frames with a source-derived full inverse and coupled local-coordinate export. Native target frames and exact saved-file rendering are tested; native intermediate UI semantics, grouped text, child rotation, shear and group authoring remain open.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## P-TABLE: Tables, charts and SmartArt

Dependencies: P-DRAW. Current entry points / proposed ownership: Editable slide table/chart/diagram models.

Acceptance scope: Create/edit tables, merge/split cells, chart data/series and diagram layouts/text; keep native relationships and compare rendered output after data changes.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## P-SLIDES: Slide management, sections and reuse

Dependencies: P-MASTER. Current entry points / proposed ownership: Slide identities and relationship graph.

Acceptance scope: Insert/reuse/duplicate/delete/reorder slides and sections across decks; preserve theme choices, links, media and animations through undo and repeated exports.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## P-ANIMATION: Animation timeline and transitions

Dependencies: P-DRAW,P-SLIDES. Current entry points / proposed ownership: Proposed timing graph and player.

Acceptance scope: Entrance/emphasis/exit/motion paths, triggers, build sequences, durations, delays and transition variants; author and play back with native timing/event comparisons.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## P-MEDIA: Audio, video and recording

Dependencies: P-ANIMATION. Current entry points / proposed ownership: Local media adapters; timeline.

Acceptance scope: Embed/link local media, trim, volume/fades, poster frames, playback triggers, captions and narration; record permission failures and unsupported codec behavior.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## P-SHOW: Slide show, presenter view and interaction

Dependencies: P-ANIMATION,P-MEDIA. Current entry points / proposed ownership: Presentation controller.

Acceptance scope: Custom shows, hidden slides, rehearsed timings, hyperlinks/actions, keyboard/pointer/ink, kiosk and presenter display; test navigation and state across multiple windows/displays where APIs permit.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## P-REVIEW: Comments, proofing and collaboration

Dependencies: P-TEXT,S-ACCESS,S-NETWORK. Current entry points / proposed ownership: Review anchors and collaboration adapter.

Acceptance scope: Threaded/review workflows and compare/merge where native version supports them; preserve authoring and anchors through slide/object edits and offline conflicts.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## P-OUTPUT: Notes, handouts, print and video export

Dependencies: P-SHOW,S-PRINT. Current entry points / proposed ownership: Notes/handout renderer; media exporter.

Acceptance scope: Speaker notes, notes/handout masters, print options, PDF/image/video outputs; compare geometry, page/slide ranges and animation timing for supported codecs.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## P-AUTOMATION: PowerPoint automation and extensions

Dependencies: P-SLIDES,S-AUTOMATION. Current entry points / proposed ownership: Presentation capability adapter.

Acceptance scope: Presentation/slide/shape object operations, action settings and extension equivalents; macros/add-ins need execution contracts, not just preserved bytes.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## P-FILES: Presentation formats and round-trip fidelity

Dependencies: P-OUTPUT,S-PACKAGE,S-PERF. Current entry points / proposed ownership: PPTX retained/new adapters.

Acceptance scope: PPTX/PPTM/PPT/POTX/POTM/PPSX/PPSM, fonts, encrypted content, embedded objects and large decks; verify every represented and opaque relationship after cross-app editing.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.
