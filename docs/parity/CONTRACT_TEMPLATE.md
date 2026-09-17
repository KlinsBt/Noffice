# FEATURE-ID: specific user capability

Copy into docs/parity/contracts/FEATURE-ID.md. Replace every placeholder before status specified; generated scaffolding is pending. One contract is a bounded behavior, not a whole tab.

## Identity and scope

- Application, parent family, registry ID, status and dependencies.
- Owned app:ControlId values and every source-row/context placement, or explicit non-ribbon workflow. Distinguish related partial commands from owned complete commands.
- Native version/edition assumptions; source URLs and retrieval dates; baseline reference.
- User-visible outcome; supported objects/formats; interaction sequence; complete options and preconditions.
- Explicit remaining differences. An alternative is labeled equivalent, not exact.

## Behavioral specification

Enumerate numbered cases with fixture, initial state, exact user action, expected document changes, selection/focus outcome and expected error/cancellation behavior. Include menu/dialog/keyboard variants, empty/mixed selections, protected/read-only content, boundary values, locale differences and applicable feature combinations. Specify how repeated edits, undo and redo work.

## Model and implementation

List source entry points, immutable transaction boundaries, stable source identities, schema migration, layout/render changes and reference repair. Specify retention of unknown nodes and dependency recalculation. Define relevant limits and cancellation; limits smaller than the native target remain gaps.

## Acceptance matrix

| Dimension   | Exact cases and expected results                            | Test command/artifact | Result  |
| ----------- | ----------------------------------------------------------- | --------------------- | ------- |
| import      | Fill in                                                     | Fill in               | pending |
| render      | Fill in                                                     | Fill in               | pending |
| create      | Fill in                                                     | Fill in               | pending |
| edit        | Fill in                                                     | Fill in               | pending |
| persist     | Undo/redo, reload, recovery and migration                   | Fill in               | pending |
| export      | Native reopen, edit and return; preserved unrelated content | Fill in               | pending |
| interaction | Shortcuts, focus, accessibility, responsive behavior        | Fill in               | pending |
| errors      | Rejection, cancellation, hostile/oversized input            | Fill in               | pending |
| performance | Workload, measured bounds, hardware/browser                 | Fill in               | pending |
| native      | Independent expected values/types/geometry/behavior         | Fill in               | pending |

Non-applicable dimensions need a specific rationale and reviewed artifact; do not mark native evidence non-applicable for an Office behavior claim. A view-only operation still needs unchanged-content evidence. Equality is exact for identities, text, structure and discrete results. Numeric/image tolerances must be chosen from the required behavior before running the comparison, with no broad masks hiding missing content.

## Evidence and completion

- [ ] Contract covers all owned placements and options; no unspecified cases remain.
- [ ] Relevant implementation and migrations complete.
- [ ] All applicable matrix cases pass on the same implementation fingerprint.
- [ ] Native and cross-browser artifacts reviewed; normalizations explicit.
- [ ] Proof JSON validates; parent family still open until all children complete.

Link the proof receipt and root TASKS.md evidence. Add an exact handoff for unfinished work.
