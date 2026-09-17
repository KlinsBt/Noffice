# Noffice implementation guidelines

- Do not initialize Git, create repositories/worktrees, or make commits unless the user explicitly requests it. Keep edits in this workspace; the user removed the earlier Git checkpoint.

- Use Svelte 5, SvelteKit, TypeScript, and `@sveltejs/adapter-static`. This is an explicit user requirement. Do not introduce React or another application framework.
- Ship a fully client-side static application. No server endpoints, document uploads, accounts, telemetry, remote fonts, or runtime CDN dependencies.
- `src/app.html` is the HTML shell; `src/routes` owns routing. `build/` is the release artifact.
- Use `reference_repo` for architectural inspiration. Its MPL-2.0 source is not part of this MIT codebase; do not copy implementations into MIT files.
- Keep document models, storage, formula evaluation, and conversion adapters independent from Svelte components.
- Never mark a capability complete just because a button exists. A task needs implementation, passing relevant tests, and documented limits.
- Keep `TASKS.md` as the active queue: update existing scope rows in place, put detailed evidence/handoffs in feature contracts, and keep historical delivery logs in `TASK_HISTORY.md`. Do not append another unchecked copy of the same requirement each session or reduce the parity denominator during consolidation.
- Never advertise desktop Microsoft Office parity without a validated feature inventory and independent Office compatibility evidence.
- Record persistent browser limitations and browser-related uncertainty as **Important** in `docs/parity/BROWSER_LIMITATIONS.md`, preserve the failed evidence and unfinished requirements, and continue the earliest actionable checklist dependency. Revisit exhausted investigations only with new evidence; application defects and missing verification are not automatically browser limitations.
- Treat imported documents as untrusted. Preserve originals, sanitize rendered content, bound inputs, and expose conversion losses.
- Resolve IndexedDB saves only after transaction completion. Reject stale revisions. Do not overwrite newer work from another tab.
- Follow `docs/STYLE_GUIDE.md`, `docs/ARCHITECTURE.md`, and `docs/TESTING.md`.
- For desktop-equivalence work, start with `docs/parity/README.md` and follow `docs/parity/LLM_WORKFLOW.md`. The installed Office builds in `docs/parity/baseline.json` are the native reference; newer features require supplemental evidence.
- Use `docs/parity/ENGINE_PLAN.md` and the family dependency register to select the next contract. Research routine unknowns on the PC and in official sources before asking the user. Preserve the active contract and exact next action across handoffs.
- Record command ownership and statuses in `docs/parity/registry.json`; run `npm.cmd run parity:generate` after mapping changes and `npm.cmd run parity:check` before checking completion. A tested subset does not certify its whole catalog command, family or tab.
- Desktop integrations should be pursued through tested browser equivalents as specified in `docs/parity/BROWSER_EQUIVALENTS.md`. Keep observable differences and missing evidence explicit; never delete them from the parity denominator.
- A full-parity release requires the matrix in `docs/parity/ACCEPTANCE.md` and `npm.cmd run parity:gate`. Passing consistency checks or historical test totals is not product completeness.
- Use `npm.cmd` on Windows when PowerShell execution policy prevents `npm.ps1`.
