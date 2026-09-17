# Reusable implementation prompt

Paste the following for each continuation. The target is near-complete functionality and compatibility within the existing static-browser requirements. A numeric parity claim requires an agreed feature-level denominator and current reviewed evidence; this prompt does not establish a percentage.

```text
Continue implementing Noffice toward near-complete desktop Word, Excel and
PowerPoint functionality and file compatibility. Treat every existing product
requirement as required until I explicitly approve a scope change.

Follow AGENTS.md, TASKS.md, docs/CODEBASE_AUDIT.md, docs/parity/README.md,
docs/parity/LLM_WORKFLOW.md, docs/parity/ENGINE_PLAN.md, the family dependency
register and the active contract's latest handoff. Inspect the actual files and
current failures before relying on earlier completion claims.

Implement the earliest unblocked foundation in dependency order. Finish coherent
end-to-end behavior: model, commands, UI, persistence, import/export and recovery.
Complete existing requirements and their dependencies; avoid repeatedly adding
tiny checked milestones while leaving the same parent work untouched. Finish one coherent existing TASKS.md requirement and its dependencies at a time.
Record persistent browser limitations and browser-related uncertainty as Important
in docs/parity/BROWSER_LIMITATIONS.md, keep the affected requirements open, and
continue the earliest actionable existing checklist work when those are the
remaining blockers. Revisit exhausted investigations only with new evidence. A bounded fix does not trigger switching to
another application or creating a replacement subset checkbox. Do not stop merely
after a plan, one small fix, or a passing build. Make as much verified progress as
the session permits.

Compare the same authored files and actions with the installed Office builds in
baseline.json. Test creation/import, rendering, real editing, undo/redo, reload,
actual exported files, native reopen/edit and reimport. Cover invalid input,
unsupported cases, storage conflicts/failures and applicable performance limits.
Use native UI evidence where COM setters miss interaction behavior. Preserve
originals, unrelated document parts and current working features. Run relevant
regressions and keep evidence bound to the actual tested implementation.

Keep Svelte 5, SvelteKit, TypeScript, adapter-static, MIT licensing and entirely
client-side/offline operation. Use the defined tested browser equivalents where
required, with observable differences recorded. Research routine unknowns using
the PC and authoritative sources. Ask only for missing access, product decisions
or approvals explicitly required by the execution environment.

Update existing TASKS.md rows and its dated summary in this workspace. Put detailed
evidence and exact handoffs in contracts; put historical logs in TASK_HISTORY.md.
Do not duplicate requirements, shrink the inventory, relax acceptance thresholds,
or check an unfinished parent. Report features completed and remaining failures;
do not infer a parity percentage from checkbox, command or test counts. Claim
full parity only after the complete acceptance matrix and strict release gates
pass with reviewed current native evidence.

Do not initialize Git, create repositories/worktrees, or make commits. Preserve
private fixtures and useful native evidence. If the session ends, leave the exact
active contract, changed files, commands/results, artifact paths, unresolved cases
and next executable action. Do not describe the whole programme as finished.
```
