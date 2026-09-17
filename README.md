# Noffice

A private browser office with a simple Word, Excel, or PowerPoint mode chooser. Built with **Svelte 5, SvelteKit, TypeScript, and `adapter-static`**. MIT licensed, with no account, backend, document upload, telemetry, or remote font dependency.

**Status: early implementation. Desktop Microsoft Word, Excel, and PowerPoint parity has not been achieved.** The app includes functional editors and partial Office format conversion; see [compatibility limits](docs/COMPATIBILITY.md) and the [task checklist at the repository root](TASKS.md).

**Task list: [TASKS.md](TASKS.md)** — current work, test evidence, next tasks, and the full remaining parity inventory.

The [every-tab command inventories](docs/ribbon/README.md) track the complete named placements in Microsoft's pinned Office 2016 Word, Excel and PowerPoint catalogs. Current Word additions include case changes, character format painting, indentation, formatting marks and View/Review controls; [scope and evidence](docs/WORD_HOME.md).

## Run locally

For continuing feature development, start with the [desktop compatibility programme](docs/parity/README.md). It pins the installed Office reference, tracks all catalog commands, defines browser equivalents and provides the LLM execution and evidence protocol.

Requires Node.js 22.20+ (Node 24 also supported).

```sh
npm install
npm run dev
```

On Windows, use `npm.cmd` if PowerShell blocks `npm.ps1`.

## Build and verify

```sh
node scripts/download-ribbon-catalog.mjs
npm run parity:test
npm run parity:check
npm run check
npm test
npm run build
npx playwright install chromium
npm run test:e2e
```

The pinned catalog download is a one-time development prerequisite; it validates source hashes and stays outside the shipped app. `npm run verify` runs parity integrity tests/checks and the application checks together (the catalog cache and browser must already be installed). `npm run parity:gate` separately rejects an incomplete desktop-parity release; its current failure is expected. The static adapter emits **`build/`**. Serve that directory on a static HTTP host; no Node.js server is needed for deployment. `npm run preview` previews the production build locally. The HTML source is `src/app.html`; the application lives in SvelteKit routes and native `.svelte` components.

The browser suite serves `build/` directly at `127.0.0.1:4183` using a small test-only static server. The verification workflow also checks that no runtime server routes, React dependencies, or JSX application files have been introduced.

## Included workflows

- Landing page with three mode choices. An Open dialog holds recent files, sample templates, search, filters, starred files, duplicate, and recoverable Trash.
- Word mode: rich text, headings, lists, alignment, colors, tables, local images, links, find/replace, and browser printing.
- Excel mode: editable cells, lookup and other supported formulas, ranges, clipboard operations, sorting, fill-down, fonts, number/date formats, multiple sheets, and chart previews. Imported forms display merges, custom dimensions, hidden content, borders, frozen panes, notes, and raster drawings; supported validation lists and locked cells govern input.
- Excel calculation includes tested absolute names, table references, INDIRECT, multiple criteria and scalar XLOOKUP/XMATCH. See the [native calculation evidence and remaining gaps](docs/EXCEL_CALCULATION.md).
- Six core Excel aggregates distinguish references, direct arguments, errors and empty text, with 128 native Excel comparisons. See [aggregate evidence and limits](docs/EXCEL_AGGREGATES.md).
- Loan/savings and depreciation calculations include ten financial functions verified against 170 native Excel cases. See [financial function scope](docs/EXCEL_FINANCIAL.md).
- Excel Format cells (Ctrl+1) offers presets, custom number/date codes and a live preview for the selected range. See [verified formatting and limits](docs/NUMBER_FORMATS.md).
- Excel filtering supports selected values, comparisons, reapply/clear and hidden-row-aware SUBTOTAL results. See [filtering, native Excel evidence and limits](docs/EXCEL_FILTERS.md).
- PowerPoint mode: slides, text, simple shapes, images, positioning/resizing/rotation, notes, and keyboard-controlled presentation mode.
- Shared font picker: bundled offline fonts, system-font choices, and local TTF/OTF/WOFF/WOFF2 import with IndexedDB persistence. See [font support and limits](docs/FONTS.md).
- Word typography includes point sizes, subscript/superscript, paragraph spacing/direction, inherited DOCX fonts and whole-document orientation. Supported imported paragraph and footnote edits preserve source package content. Spreadsheets and slides preserve supported font choices in Office exports.
- Word supports soft breaks, explicit tabs and page breaks (Ctrl+Enter) in supported imported paragraphs. Imported PowerPoint decks can reorder existing slides and flip images while retaining source parts; section-aware restructuring remains incomplete.
- IndexedDB autosave with stale-tab conflict checks and the last 20 versions of each file.
- Partial DOCX/XLSX/PPTX import and export, plus TXT/HTML/CSV and native `.noffice` backups.
- Locally cached production assets support reopening the workspace offline after the initial cache installation.

New Office imports return original bytes on unchanged export. XLSX patches supported edits into the source package. DOCX now patches supported paragraph, formatting and footnote edits, refusing structural edits it cannot preserve. PPTX patches supported existing text, formatting, geometry and notes, and adds simple objects/images while retaining source parts; unsupported imported structural edits fail explicitly. Imported slide proportions and placeholder positions are retained. See [PowerPoint preservation](docs/PPTX_PRESERVATION.md) for the tested scope. Preservation does not imply full browser editing/rendering support. See [Word preservation](docs/DOCX_PRESERVATION.md), the [price-schedule audit](docs/XLSX_FIDELITY_AUDIT.md) and [60-file interoperability corpus and results](docs/OFFICE_CORPUS.md). Reopen older imports to obtain new preservation metadata.

**Reopen the original XLSX when using a workbook imported by an older Noffice build.** Older saved models lack the source-part mapping required for preservation; export reports this instead of silently rebuilding a reduced workbook. Keep native backups before reopening.

Browser data belongs to the current origin and browser profile. Clearing site data removes local files. Use Export or the workspace backup action for durable copies.

## Project documents

- [Architecture](docs/ARCHITECTURE.md)
- [Style guide](docs/STYLE_GUIDE.md)
- [Research and primary sources](docs/RESEARCH.md)
- [Testing policy](docs/TESTING.md)
- [Excel calculation evidence](docs/EXCEL_CALCULATION.md)
- [Excel number-format editing and evidence](docs/NUMBER_FORMATS.md)
- [Excel filters, row visibility and subtotals](docs/EXCEL_FILTERS.md)
- [Downloaded Office corpus, results and desktop evidence](docs/OFFICE_CORPUS.md)
- [Task checklist](TASKS.md)
- [Compatibility matrix](docs/COMPATIBILITY.md)
- [Third-party notices](THIRD_PARTY_NOTICES.md)
- [Contribution guidelines](CONTRIBUTING.md)

`reference_repo/` is an external MPL-2.0 reference project and is not included in the Noffice build or MIT grant. Noffice uses independently written code inspired by its architecture.

Noffice is not affiliated with Microsoft. Microsoft Office, Word, Excel, and PowerPoint are Microsoft trademarks.

Recent verified scope: [Word paragraph layout](docs/WORD_PARAGRAPHS.md), [Excel scalar statistics](docs/EXCEL_STATISTICS.md), and [PowerPoint retained editing and rotation](docs/PPTX_PRESERVATION.md).
