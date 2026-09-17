# Third-party notices

September 11 continuation: the complete pinned upstream saxes 5.0.1 notice is now bundled. The September 12 PDF integration raises the inventory to 193 production packages; eight have unresolved full notices, including `@pdf-lib/fontkit`. Bundled upstream components of the WOFF2 decoder and fontkit also require provenance review. The offline `npm.cmd run notices:check` rejects missing or stale notices; three verifier tests cover failure cases. [Reviewed supplement, remaining packages and exact next actions](docs/licenses/README.md).

Noffice's own source is MIT licensed. Third-party packages retain their own licenses and copyright notices. The lockfile pins the installed dependency graph.

The development command inventories under `docs/ribbon/` derive from Microsoft Corporation's Office Fluent UI command catalogs, revision `b230a0df45036b2d5e8b49ebf368b3f90ada8d63`, under MIT. Their [copyright and license](docs/ribbon/LICENSE.microsoft.txt), [source hashes](docs/ribbon/source.json) and [reproduction instructions](docs/ribbon/README.md) accompany the data. They are documentation assets and are not included in the static application build.

| Direct dependency family                                    | License                                        | Purpose                              |
| ----------------------------------------------------------- | ---------------------------------------------- | ------------------------------------ |
| Svelte, SvelteKit, adapter-static, Vite Svelte plugin       | MIT                                            | Application framework and build      |
| Tiptap open-source extensions / ProseMirror                 | MIT                                            | Rich-text editor                     |
| Lucide Svelte                                               | ISC                                            | UI icons                             |
| docx                                                        | MIT                                            | DOCX creation                        |
| Mammoth                                                     | BSD-2-Clause                                   | DOCX semantic import                 |
| ExcelJS                                                     | MIT                                            | XLSX import/export                   |
| SSF                                                         | Apache-2.0                                     | Excel number and date format display |
| PptxGenJS                                                   | MIT                                            | PPTX creation                        |
| JSZip                                                       | MIT OR GPL-3.0-or-later (MIT option used)      | Office ZIP containers                |
| DOMPurify                                                   | Apache-2.0 OR MPL-2.0 (Apache-2.0 option used) | HTML sanitization                    |
| Inter, Source Serif 4, JetBrains Mono (Fontsource packages) | OFL-1.1                                        | Bundled offline fonts                |
| pdf-lib, @pdf-lib/fontkit, woff2-encoder | MIT declared; full fontkit/upstream component notices under review | Local searchable PDF generation and font decoding |
| Zod                                                         | MIT                                            | Native model validation              |

Development dependencies (TypeScript, Vite, Vitest, Playwright, jsdom, fake-indexeddb, Prettier) retain their package licenses. `npm run notices` generates `static/THIRD_PARTY_LICENSES.txt` and `docs/dependency-licenses.json` from installed package metadata, notice files, and explicit README license sections. The current inventory includes unresolved missing notices, including missing license metadata in the legacy `buffers` dependency. Public distribution remains gated on completing that review; a generated inventory does not establish permission for missing or unknown entries.

The external `reference_repo/` directory is MPL-2.0 and remains separately licensed. Its source is not imported by, copied into, or bundled with the MIT application.

The optional Office test corpus downloads Apache POI fixtures from a pinned upstream revision. Apache POI declares Apache-2.0; its complete upstream LICENSE and NOTICE are saved under `.local/office-corpus/licenses/`. The assets remain externally licensed and are excluded from Noffice's MIT grant and static build. The public manifest/downloader do not bundle those files. See `docs/OFFICE_CORPUS.md` for provenance and test scope.

Optional local PDF inspection uses pypdfium2 5.13.0 (Apache-2.0 / BSD-3-Clause) and the PDFium/dependency notices supplied in its wheel. It is installed only under `.local/pdf-tools`, never shipped as an application dependency. Redistribution of those test tools must retain their supplied notices.
