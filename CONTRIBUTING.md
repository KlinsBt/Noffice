# Contributing

Read `AGENTS.md` and the documents under `docs/` before making changes.

Use SvelteKit with the static adapter and native Svelte components. Keep application data on the client. New dependencies must be suitable for redistribution with an MIT application, and their license notices must be retained.

For a feature change, describe the user-visible behavior, the supported limits, and meaningful test evidence. Update the compatibility matrix and checklist. Never check a desktop-parity task merely because a basic implementation passes its own tests.

Run `npm run verify` before submitting changes. Use `npx prettier --check src tests` for source formatting. Do not include document content in diagnostic logs or commit private test files. Use synthetic or explicitly redistributable Office fixtures.

Contributions to Noffice are provided under its MIT license. Do not copy MPL source from `reference_repo` into MIT application files.
