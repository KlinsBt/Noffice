# Desktop compatibility programme

Start here for continued implementation. The user authorized selecting the installed Office baseline, researching missing behavior, updating the architecture and pursuing browser equivalents autonomously. This programme preserves the MIT, SvelteKit static, entirely client-side product requirements.

## Target and evidence

[BASELINE.md](BASELINE.md) pins the installed Office 2016 binaries. Their behavior is the primary reproducible reference; newer functions remain additional work, never silently certified by an older application. [SOURCES.md](SOURCES.md) records official research starting points.

The existing ribbon catalogs contain 11,206 placements. [COMMAND_COVERAGE.md](COMMAND_COVERAGE.md) groups repeated IDs while retaining every source row. A grouped command must cover all its contextual placements and options before verification. The ledger is a research queue, not a completed specification. A retained object or a passing self-round-trip does not complete an editing feature.

## Working documents

- [IMPLEMENTATION_PROMPT.md](IMPLEMENTATION_PROMPT.md): reusable continuation instructions for implementation, native verification, truthful completion and exact handoffs.

- [LLM_WORKFLOW.md](LLM_WORKFLOW.md): execution, research, testing, completion and handoff instructions.
- [ENGINE_PLAN.md](ENGINE_PLAN.md): ordered architecture work and concrete next tasks.
- [SHARED.md](SHARED.md), [WORD.md](WORD.md), [EXCEL.md](EXCEL.md), [POWERPOINT.md](POWERPOINT.md): 51 open capability families with dependencies and acceptance scope, including non-ribbon workflows.
- [CONTRACT_TEMPLATE.md](CONTRACT_TEMPLATE.md): required detailed specification for each implementable feature.
- [ACCEPTANCE.md](ACCEPTANCE.md): test matrix, native reference procedure and evidence format.
- [BROWSER_EQUIVALENTS.md](BROWSER_EQUIVALENTS.md): proposed implementations and measurable differences for desktop integrations.
- [BROWSER_LIMITATIONS.md](BROWSER_LIMITATIONS.md): **Important** observed browser constraints and persistent uncertainty, retained evidence gaps, and conditions for revisiting them while continuing actionable checklist work.
- [registry.json](registry.json): authoritative machine-readable family/contract statuses and command ownership.
- [contracts/X-TABLE-ENTRY-SUBSET.md](contracts/X-TABLE-ENTRY-SUBSET.md): migration example for existing bounded evidence, deliberately partial.
- [contracts/W-MODEL-SECTIONS.md](contracts/W-MODEL-SECTIONS.md): implemented Word source-section/migration foundation and remaining effective-layout acceptance, partial.

## Commands and completion

On a fresh checkout, run `npm install`, then `node scripts/download-ribbon-catalog.mjs` once. The downloader verifies the pinned source hashes and writes only the development cache under .local/ribbon-catalog. Subsequent parity checks revalidate that cache offline. No installed Office is needed for ledger validation; native feature certification uses the Windows reference separately.

Run `npm.cmd run parity:check` for inventory/registry/evidence consistency, `npm.cmd run parity:test` for validator regressions and `npm.cmd run parity:generate` after assigning commands. `npm.cmd run parity:gate` additionally fails while any command or family lacks verified exact behavior. Its expected current failure is not an infrastructure error. `npm.cmd run parity:fingerprint` prints the hashes needed by evidence receipts without changing the ledger.

No percentage is currently certified. A proof receipt checks integrity and traceability; it cannot establish that a test assertion is correct or a specification exhaustive. Review actual native artifacts and coverage. Browser equivalents with observable differences remain distinct from exact desktop parity. No task may be deleted or relabeled as equivalent merely to improve completion numbers.

## Programme setup checklist

- [x] Inspect installed executable versions, architecture, locale and SHA-256 without opening user documents; recorded in baseline.json.
- [x] Define family-level gaps, dependencies, browser alternatives and a contract/evidence protocol.
- [x] Verify the ledger and evidence-validator infrastructure: 24 regression tests and parity:check pass; the strict gate rejects the incomplete product.
- [ ] Research and specify every command, menu option, contextual placement and non-ribbon workflow.
- [ ] Implement and independently verify every specified capability.
- [ ] Close all critical correctness, corruption and unsupported-workflow gaps.
- [ ] Pass the complete release matrix and strict parity gate on one reproducible build.

The first two checkmarks record documentation work and a local inspection, not implementation of the office engines. Validator results are recorded in TASKS.md after execution.
