# Excel capability contracts and gap register

The reference foundation retains lazy row/column shape and tested SUMPRODUCT arithmetic, constants, coercion/error ordering, names/table columns and INDEX/INDIRECT consumers. Single-cell CSE intent and the measured IF/IFERROR/IFNA array consumers now have native entry/history/reopen and actual-export evidence; all 47 initial mismatches are resolved across 96 native snapshots. Array metadata hydrates safely from retained source and unsupported multi-cell member edits reject. All 39 earlier scalar-reference and 48 shaped-reference cases remain tested. Full conditional/function/type/locale/laziness matrices, multi-cell CSE, spills, fresh-workbook native exports, complete table selectors, static names and workers remain required. [Current contract](contracts/X-MODEL-DEPENDENCIES.md#single-cell-cse-identity-and-conditional-array-continuation).

Every family below is an open umbrella acceptance gate. Existing subsets do not complete the family. Split it into concrete contracts using [CONTRACT_TEMPLATE.md](CONTRACT_TEMPLATE.md), and assign every relevant command from [COMMAND_COVERAGE.md](COMMAND_COVERAGE.md). Non-ribbon behavior listed here must also receive contracts. These are implementation requirements, not a claim that detailed behavior has already been researched.

## X-CALC: Formula grammar, functions and dependency graph

Dependencies: S-MODEL. Current entry points / proposed ownership: src/formulas.ts; proposed worker calculation service.

Acceptance scope: Inventory every installed function and operator; test type coercion, blanks/errors, lazy branches, references, dates, names, locales, precision, volatile and iterative calculation; use native inputs/outputs with stated tolerances.

The bounded [static-reference/blank-cell subset](contracts/X-MODEL-DEPENDENCIES.md#static-reference-and-blank-cell-subset) adds compact cell/range identities, axis anchors and conservative retained-export invalidation alongside observed edges. Native fixtures distinguish scalar blanks, empty-text formulas, inactive branches and the tested blank consumers. The [relative-name subset](contracts/X-MODEL-DEPENDENCIES.md#relative-name-anchor-subset-active-september-11-2026) now binds imported A1-origin names at each caller, including mixed axes, aliases/local shadows and tested range consumers. Complete reference/value types, implicit intersection, arrays, static name bindings, workers and the full family matrix remain open.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-ARRAY: Legacy arrays, dynamic arrays and modern formulas

Dependencies: X-CALC. Current entry points / proposed ownership: Array-valued formula model.

Acceptance scope: Legacy CSE arrays, spill ranges, implicit intersection, LET/LAMBDA and modern functions; use installed Excel where supported and authoritative newer oracles otherwise. Unsupported native functions are not a passing test.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-GRID: Worksheet structure and editing

Dependencies: X-CALC,S-CLIPBOARD. Current entry points / proposed ownership: SheetEditor.svelte; sheet-structure.ts; workbook checkpoints.

Acceptance scope: Insert/delete/move/copy rows, columns, cells and sheets; imported sheet rename/reorder; repair all references, rules, drawings and names. Fill series, Flash Fill, auto-fit and cell selection need separate contracts.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-FORMAT: Cell formatting, themes and rich text

Dependencies: X-GRID,S-TEXT. Current entry points / proposed ownership: Number formats; styles; font adapters.

Acceptance scope: Full Format Cells options, rich runs, borders, custom formats, locale/date systems, row heights and theme inheritance; compare stored types, width-sensitive display and print output.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-TABLE: Structured tables and calculated columns

Dependencies: X-GRID,X-FORMAT. Current entry points / proposed ownership: sheet-tables.ts; sheet-table-entry.ts; xlsx-tables.ts.

Acceptance scope: Complete table lifecycle, styles, headers, totals, column operations, row insertion, direct/paste/horizontal/Tab growth, formula propagation and preferences; preserve reference semantics across every combination.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-FILTER: Sorting, filtering and outlines

Dependencies: X-TABLE. Current entry points / proposed ownership: sheet-filters.ts; outline/group model.

Acceptance scope: Multilevel stable sort, custom lists, color/icon/date/top filters, slicers, hidden/filtered distinctions, subtotals and grouping; verify visible sets, rule XML and recalculation after edits.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-VALIDATE: Validation, protection and names

Dependencies: X-CALC,X-GRID. Current entry points / proposed ownership: sheet-validation.ts; name/protection editors.

Acceptance scope: All rule/alert kinds, circles, dynamic lists, scoped/relative names, worksheet/workbook permissions and passwords; test editable cells under protection without bypassing the source lock. Keep all-locked corpus case as a protection test.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-CONDITIONAL: Conditional formatting rules and visuals

Dependencies: X-CALC,X-FORMAT. Current entry points / proposed ownership: Proposed rule evaluator and renderer.

Acceptance scope: Rule priority, Stop If True, relative formulas, duplicate/top/date rules, data bars, scales and icon sets; range edits repair rules and match displayed native outcomes.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-CHART: Charts, sparklines and drawing editing

Dependencies: X-CALC,S-DRAW. Current entry points / proposed ownership: Chart data/series model; sheet drawings.

Acceptance scope: Chart types, axes, trends, error bars, labels, styles, combo/3D charts and source edits; all drawings editable; compare chart XML, data and native rendered output.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-PIVOT: Pivot tables, pivot charts and caches

Dependencies: X-CALC,X-FILTER,X-CHART. Current entry points / proposed ownership: Proposed pivot engine and cache adapter.

Acceptance scope: Create/refresh pivots, aggregation, grouping, calculated fields/items, sorting/filtering, layouts, drill-down and slicers; verify native results and cache consistency after source changes.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-DATA: Connections, queries and data model

Dependencies: X-PIVOT,S-NETWORK. Current entry points / proposed ownership: Proposed local query and relational model workers.

Acceptance scope: Import transformations, refresh, relationships, query language and analytical expressions; cover Power Query/Power Pivot capabilities with version/edition evidence and explicit local-source equivalents.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-ANALYSIS: What-if, Solver and analysis tools

Dependencies: X-CALC,X-DATA. Current entry points / proposed ownership: Proposed numerical/analysis engines.

Acceptance scope: Goal Seek, scenarios, data tables, optimization and analysis outputs; define convergence, numeric tolerances, infeasible/error cases and reproducibility against installed tools where available.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-VIEW: Views, review and print layout

Dependencies: X-FORMAT,S-PRINT,S-ACCESS. Current entry points / proposed ownership: Pane/view commands; comments; print settings.

Acceptance scope: Split/freeze, custom views, page breaks, scaling, repeated titles, headers/footers, comments/review and accessibility; compare printed pages and keyboard navigation.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-AUTOMATION: Excel VBA, add-ins and controls

Dependencies: X-GRID,S-AUTOMATION. Current entry points / proposed ownership: Workbook automation adapter.

Acceptance scope: Macro recording, events, UDFs, workbook object operations and form controls; execute supported scripts deterministically offline; preserve unsupported macro binaries and disclose execution differences.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.

## X-FILES: Excel formats and high-volume workbooks

Dependencies: X-PIVOT,X-VIEW,S-PACKAGE,S-PERF. Current entry points / proposed ownership: Workbook format adapters.

Acceptance scope: XLSX/XLSM/XLSB/XLS/XLTX/XLTM and CSV encodings, protected/encrypted files, external links and large sparse sheets; verify editable content and numeric results, not merely successful ZIP creation.

- [ ] Research native options, states, shortcuts, errors and combinations; pin sources and reproducible fixture steps.
- [ ] Decompose into command and non-ribbon contracts with exact expected results and explicit tolerances.
- [ ] Implement model, UI, import, rendering, export and migrations as applicable.
- [ ] Pass undo/reload, native comparison, cross-browser, accessibility and stress/error cases.
- [ ] Attach current hashed evidence and verify all child contracts before completing this umbrella gate.
