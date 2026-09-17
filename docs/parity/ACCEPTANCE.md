# Acceptance and proof protocol

## Native comparison loop

Create/read a fixture in the pinned Office app, snapshot its initial state, perform the specified UI actions, and save a native expected copy. Import the same fixture in the static browser app, perform the same actions, download the actual result, reopen it in Office, edit again and return it to Noffice. Compare content/model values and types, geometry, semantics and rendering. Hash original/native/browser artifacts and the test code. Inspect unknown package parts independently of editable features.

For Word compare run properties, sections, fields/revisions, line/page positions, page counts and rendered PDFs. For Excel compare values and types, formulas and errors, number formats, names/ranges, dependency changes, tables/pivots/charts, visibility and printed pages. For PowerPoint compare object hierarchy, source identities, geometry, rich runs, inheritance, links, notes, playback events/timing and rendered slides. A render-only screenshot does not prove editable behavior.

Record actual app Version/Build, executable baseline hash, locale, calculation/autocorrect settings, installed fonts, paper/printer settings, browser version and hardware for applicable cases. Existing scripts under scripts/verify-* and docs/TESTING.md provide tested patterns; reuse their guards. Originals and private files must remain unchanged. COM comparisons supplement real UI tests, especially for entry/autocorrect and selection semantics.

## Release matrix

- [ ] Unit tests and Svelte diagnostics pass: npm.cmd run check; npm.cmd run test.
- [ ] Production static architecture passes: npm.cmd run build.
- [ ] Chromium workflows including private fixtures pass when locally available: npm.cmd run test:e2e.
- [ ] Firefox, WebKit and actual-platform permission/printing checks are implemented and pass; current Chromium coverage is not substituted.
- [ ] External corpus workflows pass and all strict evidence gaps have an appropriate resolved test: npm.cmd run corpus:test; npm.cmd run corpus:gate.
- [ ] Native comparisons pass for all claimed feature contracts, including combinations and complete UI/dialog behavior.
- [ ] Storage failure, offline, accessibility, performance and untrusted-content suites pass on documented workloads.
- [ ] npm.cmd run parity:test and parity:check pass; parity:gate passes only for fully verified exact parity.

The protected all-locked fixture must remain locked. Add a separate authenticated/appropriately editable fixture for edit evidence rather than bypassing protection. The deeply nested DOCX gap needs bounded import/layout support and positive tests; retaining a safe rejection is not desktop functionality.

## Proof schema

Distribution notice gate: run `npm.cmd run notices:test` and `npm.cmd run notices:check`. The latter deliberately fails while missing/unknown package notices remain; generated notices alone do not complete the transitive review. See [the review register](../licenses/README.md).

Registry verified entries reference a JSON file with schemaVersion 1, contract (matching ID), result passed, baselineHash, inventoryHash, implementationHash and contractHash. The first three current hashes are printed by parity:fingerprint; contractHash is SHA-256 of the exact Markdown bytes. Include checks for all ten dimensions from CONTRACT_TEMPLATE.md. Every check records dimension, command, exitCode 0 and nonempty artifacts with workspace-relative path and SHA-256. A non-applicable check additionally sets notApplicable true and a specific reason, with a hashed rationale artifact. Store native configuration and mismatch counts in those artifacts. The validator does not execute arbitrary commands from receipts.

Example check: {"dimension":"native","command":"powershell.exe -NoProfile -File scripts/verify-feature.ps1","exitCode":0,"artifacts":[{"path":".local/feature/native-report.json","sha256":"replace with actual SHA-256"}]}.

Proofs bind src, tests, scripts, corpus, package/lock files and core build/test configuration. A shared-code change invalidates the conservative implementation hash; rerun affected feature tests and refresh receipts after evaluating the dependency impact. No fabricated receipts, stale artifact reuse or copied success values. Integrity validation cannot prove test adequacy; contract/native artifact review is mandatory. Registry status may remain partial even when existing bounded tests pass.

An exact family gate additionally needs every child contract complete and an explicit reviewed non-ribbon coverage record. The machine gate checks declared records; discovering omitted behavior must add new records and invalidate the family claim. Browser-equivalent contracts can be verified as equivalents, but do not turn the exact-parity release gate green.
