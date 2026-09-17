# Codebase audit and continuation handoff

Latest implementation, September 11: PowerPoint rotated-child quadrant normalization and numeric size-field corner retention resolve all eight group-frame differences and the separate slide-3 sequential offset difference. Native angle discovery preceded implementation. Editing/history/reload, invalid/unsupported recovery, legacy migration, actual exports/reimports and native re-edit pass. Final checks: 1,746 unit passes / one skip, 110 Chromium workflows, 63 corpus scenarios, zero diagnostics, 34 static files, 13 native PowerPoint groups and 46 exact native PDF comparisons. The pixel negative control fails correctly. Root TASKS.md is updated in place (24 checked / 44 open); strict parity/corpus/notices remain failures. The next Word mixed-hard-break probe exposes 48 paragraph-relative glyph and 24 second-line baseline differences; it is diagnostic, not completed support. [Exact implementation and handoff](parity/contracts/P-MODEL-INHERITANCE.md#quadrant-bounds-and-numeric-size-continuation).

Audited September 10, 2026 against the local workspace and installed Office executables. This is a repository-wide structural review with targeted inspection of the editing, conversion, storage, calculation, rendering, build and acceptance paths. It is not a claim that every branch or Office workflow has been manually exercised.

## September 11 workspace cleanup

The filesystem inventory covered **20,598 files**, including dependencies and the separate reference tree. The runtime graph inspected all **178 source files**, with no unreachable non-test modules or unresolved relative imports from the SvelteKit routes/service worker. All direct production dependencies have source or asset uses; the three Fontsource packages supply six existing local font files. Maintained source/tests/scripts/corpus/docs contain no exact duplicate files. This is structural analysis and targeted review, not proof that every function, CSS selector or feature is necessary/correct.

Removed **1,389 disposable files, 303,706,016 bytes**: `.local/npm-cache`, `.svelte-kit` build intermediates, `.local/ribbon-generator-test` generated comparison output, root `debug.log`, and the empty `public/` directory. `npm.cmd run check` regenerated 27,165 bytes of required SvelteKit metadata. Installed dependencies, the 34-file release build, external architectural reference, native tools, fixtures and comparison evidence remain available. Hash comparison confirms **1,880 application/fixture/release/retained-evidence files unchanged**. The implementation fingerprint remains `09c99ec4d69a4109b0811b4c01a9b70d3e4b19185d22b0b11a06d997aad6b3da`.

Automatic approval review initially rejected the broader deletion of `.gitignore`, `.gitattributes` and six one-time `.local` helpers. The user subsequently explicitly approved the two root Git configuration deletions; both are now removed. The six helpers remain retained: `check-checkpoint-bytes.py`, `finish-relative-docs.py`, `record-relative-handoff.py`, `update-empty-docs.mjs`, `update-grid.py`, `update-relative-docs.py`. `.ignore` supplies workspace search exclusions without requiring Git. AGENTS.md explicitly prohibits recreating Git or making commits without a request; the old checkpoint instruction is superseded.

Post-cleanup checks: **1,685 unit passes / one skip**, **zero Svelte errors/warnings**, `parity:check` passes, all **three notice-verifier tests** pass, and `check-static.mjs` validates the preserved **34-file build**. Browser/native suites were not rerun for this cache/documentation cleanup. No feature checkbox was changed and no new parity claim is made. Detailed inventory, removal, preservation and verification receipts use `.local/cleanup-*.json`; command logs use `.local/cleanup-*.log`.

Reusable instructions are in [IMPLEMENTATION_PROMPT.md](parity/IMPLEMENTATION_PROMPT.md). The bounded PowerPoint group-bound and Excel scalar-intersection continuations are now implemented; ENGINE_PLAN.md resumes Word coordinates/heterogeneous flow. The cleanup receipt records the two approved configuration removals and six retained helpers. Continue implementation from the active task queue. No Git repository exists.

## Earlier implementation updates

Previous Word continuation (September 11, 2026): nine explicit empty hard-break/paragraph-mark cases and terminal typing now have native evidence; Redo no longer substitutes a pasted font. All 1,685 unit tests, 102 browser workflows, 63 corpus scenarios and 16 native Word groups pass; 92 native-export PDF pages match exactly. The root checklist remains 24 checked / 44 open because no unfinished parent was certified. Missing notices decreased from eight to seven; the notice gate and three failure-path tests are implemented. Changes remain uncommitted. Exact coordinates, heterogeneous wrapping and full parity remain open; next independent foundation is PowerPoint group-bound normalization. [Current evidence and exact handoff](parity/contracts/W-TEXT-UNIFORM-LINES.md#empty-line-reproduction-and-current-handoff).

The latest Excel continuation (September 11, 2026) implements imported A1-origin relative-name binding at each caller, including mixed axes, scoped aliases, tested aggregate ranges and safe source-origin migration. Root TASKS.md is **24 checked milestones / 44 open requirements**. All 101 browser workflows, 1,683 unit tests, 63 corpus scenarios and nine native Excel groups pass; three actual exports match 147 independent native cell snapshots. All 15 native Word and ten native PowerPoint regression groups also pass on the final exports; their documented coordinate/text/group limits remain open. Cold-cache bootstrap was rechecked and the application is preserved in a local Git checkpoint with exact-byte attributes. Full name/array/worker semantics and all parent contracts remain open; next is Word secondary coordinates and empty hard-break metrics. [Current evidence and handoff](parity/contracts/X-MODEL-DEPENDENCIES.md#current-handoff).

The previous PowerPoint continuation (September 11, 2026) adds full orthogonal group inverses and coupled child-coordinate export for the measured rotated/reflected frames. Root TASKS.md is **23 checked milestones / 44 open requirements**. All 99 browser workflows, 1,679 unit tests, 63 corpus scenarios and ten native PowerPoint groups pass; ten exported slides match independent native rendering exactly. Complete group selection/bound normalization, child rotation, shear and text remain open. The next independent dependency is Excel relative-name anchors. [Current evidence and handoff](parity/contracts/P-MODEL-INHERITANCE.md#current-handoff).

The previous Excel continuation (September 11, 2026) implements compact literal static-reference identities/export invalidation and fixes the tested scalar blank-reference distinction. Root TASKS.md now has **22 checked milestones / 44 open requirements**. All 98 browser workflows, 1,673 unit tests, 63 corpus scenarios and eight native Excel groups pass. Three final exports match 84 native cell snapshots. Complete reference/name/array semantics and workers remain open; the queue now moves to PowerPoint rotated/reflected group-child edits. [Current evidence and exact handoff](parity/contracts/X-MODEL-DEPENDENCIES.md#current-handoff).

The previous Word continuation (verified September 11, 2026) corrects source text-area width, hanging wrap-spaces and eligible homogeneous wrapped-line leading. Native line identities, controlled first-line placement and three actual exports pass; secondary coordinates and heterogeneous flow remain open. Root TASKS.md has **20 checked milestones / 44 open requirements**. All 97 browser workflows, shared unit/corpus checks and 15 native Word groups pass. The queue now rotates to Excel static cell/range identities before PowerPoint group-transform editing, preserving Word's unfinished coordinate and pagination requirements. [Current evidence and exact handoff](parity/contracts/W-TEXT-UNIFORM-LINES.md#current-handoff).

The earlier scoped-name continuation implements imported Excel name scopes, workbook-alias binding, internal `[0]!Name` lookup and legacy metadata migration, with native edit/delete evidence and seven passing Excel regression suites. That continuation brought TASKS.md to **16 checked milestones / 44 open requirements**. Complete name authoring, static reference graphs, arrays and workers remain open. [Current evidence and next Word dependency](parity/contracts/X-MODEL-DEPENDENCIES.md#current-handoff).

The subsequent implementation fixed the missing CI catalog bootstrap and verified a fresh pinned download. Word's prior 8-160pt restriction is removed within its 1-1638pt UI range; native Grow/Shrink steps and the two uniform/empty-line metric cases now have targeted tests. Excel now has observed dependency tracking and immutable-revision invalidation, retaining its existing evaluator; full static binding/array/worker functionality remains open. See the updated [active checklist](../TASKS.md) and its linked contracts for current evidence. Findings below describe the audited baseline and are retained as historical context, not assertions that those specific fixes are still absent.

PowerPoint now implements the measured inherited placeholder fonts/geometry, retained rich-run display/replacement, legacy migration and first-notes master/theme package creation. The next iteration adds measured nested group frames, positive axis-aligned child move/resize and legacy group-coordinate migration. Native regressions and exact exported-slide PDF comparisons pass; arbitrary group editing, master authoring and full text layout remain open. [Current implementation and handoff](parity/contracts/P-MODEL-INHERITANCE.md#current-handoff). That earlier continuation brought the root checklist to **14 checked milestones / 44 open requirements**; the three-checkmark count below describes the original audit baseline.

## Task-file answer

There is one active root [TASKS.md](../TASKS.md). A filename scan including ignored workspace directories, excluding dependency trees, found no second copy. [TASK_HISTORY.md](../TASK_HISTORY.md) retains historical deliveries and granular requirements. The generated [command inventories](ribbon/README.md), [engine plan](parity/ENGINE_PLAN.md), family register and feature contracts contain supporting checklists with different purposes.

The preceding delivery was already recorded in the active file: omitted 10pt Word defaults, separate paragraph-mark fonts, legacy font migration and empty-paragraph typing. Those changes expanded existing checked subset rows, leaving three checked and 44 unchecked root rows. The active file now has a visible dated update explaining that. An unchecked umbrella can contain working subsets; neither that checkbox count nor the unit-test count measures parity.

The registry reports 5,579 command identities, 11,206 catalog placements, two registered contracts, 5,579 unassigned command identities and 51 open families. Unassigned does not mean that no corresponding behavior exists; it means no complete command contract owns and certifies it. Catalog identities also include containers and repeated placements, so they are not a sufficient feature denominator by themselves. Non-ribbon behavior and command options require their own coverage.

## Prioritized findings

### 1. Fresh-checkout CI lacks a required catalog input — high priority

[verify.yml](../.github/workflows/verify.yml) invokes `npm run verify` without downloading the pinned catalog. [parity-audit.mjs](../scripts/parity-audit.mjs) invokes [build-ribbon-checklists.mjs](../scripts/build-ribbon-checklists.mjs), which reads `.local/ribbon-catalog/receipt.json`. `.local/` is ignored.

Reproduction: executing the checklist builder with `--check` from a new empty directory fails with `ENOENT` for that receipt before validating any catalog. Existing dependencies were resolved from the real script location; the working catalog cache was not moved or deleted. This confirms the absent-input failure underlying the workflow issue; the complete GitHub-hosted job was not run.

Next repair: add the existing pinned downloader before verification, with its source-integrity checks intact. Prove the cold-cache bootstrap in an isolated directory and ensure ordinary verification still passes. Keep developer verification distinct from the strict release qualification job; do not make an always-failing parity gate appear to be a normal green test.

### 2. Word still changes run sizes and lacks general line metrics — high priority

[docx-typography.ts](../src/docx-typography.ts) clamps imported run sizes to 8–160pt. Calling the actual transform/decorate functions with 6pt and 200pt runs produces 8pt and 160pt respectively. The import path in [docx-import.ts](../src/docx-import.ts) uses this decorator. This changes the editable rendering model; preservation of the original DOCX does not repair its display. This limitation was already disclosed in the active contract and remains unresolved.

The [current Word handoff](parity/contracts/W-MODEL-SECTIONS.md#current-handoff-font-defaults-and-empty-paragraph-entry) separately records automatic/minimum line-box mismatches. The native empty 30pt paragraph mark with an 18pt minimum needs a 34.5pt advance in the measured fixture, while the browser uses 18pt. Measured 10pt text at 1.5-line spacing needs 17.25pt rather than 15pt. Earlier native snapshots establish these cases; this audit did not rerun Office document rendering. Exact-spacing tests and correct exported font values do not establish browser layout equivalence.

Next repair: a native-backed run-size boundary regression, then the already-authored minimum/automatic line-height cases. Derive metrics from the actual font and layout context; do not apply the measured Arial ratio universally. Complete the bounded font/line-metrics dependency before enabling variable pagination.

### 3. Excel has a calculator, but not the planned calculation architecture — high priority

[formulas.ts](../src/formulas.ts) already parses formulas into nodes and memoizes results and syntax within a calculator instance. It should be extended, not discarded on the mistaken assumption that there is no parser. [SheetEditor.svelte](../src/lib/SheetEditor.svelte) derives a new calculator from `content.sheets`; there is no persistent dependency invalidation graph or dedicated calculation worker in the inspected runtime. The service worker caches the application and is not a calculation worker.

The editor also rejects pastes beyond 10,000 rows or 256 columns and caps its add-row path at 10,000 rows. Several imported structural edits are explicitly rejected when dependent tables/charts/pivots or unsupported relationships are involved. These guards prevent incorrect edits but leave desktop workflows unavailable.

Next contract: `X-MODEL-DEPENDENCIES`. Establish stable scoped references, typed results, dependency edges and dirty propagation through edit/delete/undo/reload; compare with Excel after recalculation. Follow with worker scheduling, stale-result rejection, cancellation and measured large-workbook behavior. Raising bounds alone is not the implementation.

### 4. PowerPoint's editable model flattens rich text and shapes — high priority

[pptx-import.ts](../src/pptx-import.ts) joins text runs into paragraph strings and takes the first run's properties for an element. [model.ts](../src/model.ts) represents slide elements as text, rectangle, ellipse or image with one text-style set; [SlidePreview.svelte](../src/lib/SlidePreview.svelte) renders that single style. A placeholder containing mixed fonts or emphasis cannot be faithfully represented and edited by this model. Other preset shapes can become rectangles in the editable preview. Source-package retention is a separate capability.

[pptx-preserve.ts](../src/pptx-preserve.ts) explicitly rejects edits involving text fields/soft breaks and some imported object changes. Some layout/master lookup already exists, but it does not provide the full effective/direct inheritance and rich group model required by the pending contract.

Next contract: `P-MODEL-INHERITANCE`, using native mixed-run placeholders, two masters/layouts and nested groups. Preserve source identities, distinguish inherited from direct properties, represent runs and composed transforms, and test edits plus saved rendering and relationships.

### 5. The work queue allows one broad contract to absorb progress indefinitely — high priority

`W-MODEL-SECTIONS` includes enough typography, pagination, structural editing and acceptance work to keep it partial across many sessions. Meanwhile, the Excel and PowerPoint foundation contracts remain pending. The workflow already permits bounded child contracts, but the current queue does not make those milestones visible enough.

Proceed by completing the active, failing line-metrics dependency as a bounded child scope with independent evidence; retain the Word parent as partial. Then select the earliest unblocked Excel and PowerPoint foundation scopes using concrete prerequisite contracts. Do not require the whole Word umbrella to be verified before starting unrelated foundations. Do not reset evidence, duplicate unchecked requirements or call smaller milestones whole-command parity.

### 6. Passing development checks does not qualify a release — high priority

The current Playwright configuration runs Chromium only. The standard `verify` chain omits the strict corpus and parity gates and does not establish Windows-native UI acceptance, Firefox/WebKit, accessibility or full recovery/performance coverage. The current corpus gate fails on the deeply nested DOCX and protected-workbook editing evidence gaps. The protection case requires an appropriate editable fixture; bypassing protection is not a fix.

The [dependency inventory](dependency-licenses.json) has eight entries without bundled notice files: `binary`, `buffers`, `chainsaw`, `dingbat-to-unicode`, `https`, `is-reference`, `locate-character` and `saxes`. `buffers` additionally has unknown license metadata. This is an inventory finding, not a determination that all eight are unlicensed. Resolve provenance and the necessary notices before distribution, as the existing notices document requires.

The user explicitly removed the local Git repository. Its absence is intentional, not an issue to fix. Do not initialize Git, create repositories/worktrees or make commits unless explicitly requested. Preserve current workspace files, native evidence and private fixtures without recreating the removed checkpoint. Historical checkpoint references below describe an earlier state.

## Unused code and cleanup assessment

The structural scan covered 162 files under `src`, plus the scripts, tests, documentation, configuration and reference tree. The latter categories were surveyed by inventory and targeted reading, not line-by-line certification. A TypeScript import graph starting at SvelteKit routes and the service worker found no unreachable non-test application modules or unresolved relative source imports. This is file-level reachability, including type imports; it does not prove that every exported function or branch is used.

All direct production dependencies have source or asset uses. The three apparent Fontsource exceptions in the initial scan are referenced by `url()` in `src/font-assets.css`; they are required offline font assets. Do not remove them.

Reasonable cleanup candidates:

- `public/` is empty; SvelteKit serves `static/`. Removing an empty leftover directory has no functional benefit.
- `build/`, `.svelte-kit/`, `node_modules/`, logs and test output are generated artifacts, not duplicate implementations. Rebuild or archive them when appropriate. `.local/` additionally contains valuable native evidence and private fixtures; do not blanket-delete it.
- `styles.css` is about 3,048 lines, and ribbon/layout rules also live in separately imported stylesheets. Consolidate ownership and specificity only with viewport/print checks. Overlapping selectors are a maintenance concern, not evidence that every earlier rule is dead.
- The three editor components are roughly 1,050–1,524 lines each, with command/state logic mixed into UI. Extract commands and semantic model transitions while implementing the planned contracts, preserving undo and save behavior. A wholesale cosmetic rewrite would consume time without closing parity gaps.
- `reference_repo/` is separately licensed reference material, not part of the runtime graph. Keep it separate from MIT application source. Its size is not shipped application size.

The storage transaction-completion checks, stale-revision rejection, migrations, original-package preservation, sanitization, input bounds and explicit unsupported-edit errors are useful safeguards. Do not delete them to make features appear unrestricted.

No application code or dependency was deleted during this audit. The source import scan did not identify a defensible runtime removal.

## Validation performed in this audit

| Check | Result | Scope |
| --- | --- | --- |
| `npm.cmd run check` | Pass: zero errors/warnings | Svelte/TypeScript diagnostics |
| `npm.cmd run test` | Pass: 1,606 tests; one skipped | Current unit suites |
| `npm.cmd run build` | Pass: static architecture, 33 deployable files | Production build, not full UI behavior |
| `npm.cmd run parity:test` | Pass: 24 tests | Acceptance-ledger validation rules |
| `npm.cmd run parity:check` | Pass | Ledger/catalog consistency, not completeness |
| `npm.cmd run corpus:gate` | Fails | Existing corpus receipt: two unresolved qualification cases |
| `npm.cmd run parity:gate` | Fails | 5,630 unresolved command/family gates |
| Cold-cache catalog probe | Reproduced exit 1 / missing receipt | Local reproduction of CI prerequisite failure |
| Run-size decorator probe | Reproduced 6 → 8pt and 200 → 160pt | Actual import helper; not a new full native UI comparison |
| Installed executable SHA-256 checks | All match baseline | Word/Excel/PowerPoint binaries unchanged |

Pinned native references verified here: Word 16.0.4627.1000, Excel 16.0.4627.1001 and PowerPoint 16.0.4266.1001, all x64 Office 2016. Do not silently substitute current Microsoft 365 behavior for the apps on this PC. Additional newer behavior needs supplemental evidence.

Local audit probes and results: `.local/codebase-audit.mjs`, `.local/codebase-font-probe.mjs` and `.local/codebase-audit.json`. Reproduce them in that order; the font probe adds the CSS asset review and run-size results to the report. The cold-cache directory is `.local/audit-empty-checkout/`. Existing browser/native test receipts remain under their documented contract paths. The previous 85-workflow Chromium and 63-scenario corpus runs were not repeated during this documentation audit and are not presented as fresh runs.

## Exact handoff and suggested command

Only `TASKS.md` and this report were changed as project documentation, plus ignored local audit probes/results and regenerated build output. Runtime source, dependency versions, registry ownership and feature completion statuses remain unchanged. No missing access or product decision was encountered. Next executable repair: cold-cache CI bootstrap, validated without deleting the existing `.local` evidence. Then continue the Word font/line-metrics dependency from its linked current handoff and progress to the pending Excel and PowerPoint model contracts.

The user has already provided enough authorization for implementation. A stronger prompt is not the missing ingredient; progress requires bounded engineering work and independently checked behavior. The following command makes the order and reporting explicit:

```text
Continue implementing Noffice. Read AGENTS.md, TASKS.md,
docs/CODEBASE_AUDIT.md, docs/parity/LLM_WORKFLOW.md and
docs/parity/ENGINE_PLAN.md. Fix the audit's cold-cache CI failure first,
then the Word run-font bounds and active line-metrics dependency.
Use bounded child contracts; keep the Word parent partial and continue
to the earliest unblocked Excel and PowerPoint foundation contracts.

Use the installed Office builds in docs/parity/baseline.json as the
reference. Research routine decisions yourself. Preserve the static,
fully client-side SvelteKit architecture, existing user files, source
packages and migration/history guarantees. Implement real editable
behavior, not placeholder controls or preservation-only substitutes.

For each contract, test native behavior, browser editing, undo/redo,
reload, errors and actual exported files. Update the existing TASKS.md
rows and its dated progress summary with passing evidence; do not add
duplicate requirements or mark broader scopes complete. Continue to
the next dependency after each completed scope. Ask only for genuinely
missing access or a decision that changes product requirements.

Claim full parity only after the complete acceptance matrix and strict
release gates pass with reviewed native evidence. If the session ends
first, leave an exact handoff naming the remaining mismatch, changed
files, test commands/results, artifact paths and next executable step.
```

Browser equivalents must remain explicitly distinguished from exact Windows behavior, following `BROWSER_EQUIVALENTS.md`. If an exact requirement ultimately needs capabilities outside the authorized browser-only product, report the concrete incompatibility and available alternatives; do not silently drop it, fabricate a percentage or weaken acceptance to turn the gate green.
