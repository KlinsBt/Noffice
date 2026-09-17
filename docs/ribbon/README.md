# Ribbon command inventories

The [desktop compatibility programme](../parity/README.md) now uses these catalogs as its command-placement baseline. [COMMAND_COVERAGE.md](../parity/COMMAND_COVERAGE.md) retains all 11,206 placements under 5,579 application-specific command IDs and links assigned acceptance contracts. The machine gate rejects checked catalog rows without a verified owning contract. Unassigned commands remain pending; this deduplication does not discard contextual behavior.

These checklists cover the Office 2016 catalog baseline corresponding to the desktop-style ribbon requested by the user. The supplied screenshot shows Word's German Home tab (Start), with File, Insert, Design, Layout, References, Mailings, Review and View available across the top.

| Application | Complete catalog placement inventory       | Tab/context categories |
| ----------- | ------------------------------------------ | ---------------------- |
| Word        | [4,200 placements](WORD_COMMANDS.md)       | 49                     |
| Excel       | [3,516 placements](EXCEL_COMMANDS.md)      | 51                     |
| PowerPoint  | [3,490 placements](POWERPOINT_COMMANDS.md) | 52                     |

The data comes from [Microsoft's Office Fluent UI command identifiers](https://github.com/OfficeDev/office-fluent-ui-command-identifiers), pinned at `b230a0df45036b2d5e8b49ebf368b3f90ada8d63`. [source.json](source.json) records download URLs, Git blob hashes, SHA-256 hashes and byte counts; the [Microsoft MIT license](LICENSE.microsoft.txt) accompanies the derived catalogs. Original XLSX catalogs remain in ignored `.local/ribbon-catalog/`, outside the browser build. They are development reference data, not runtime services or application code.

A placement includes its group and parent-menu path. Duplicates occur when a command appears in several places; tabs, containers, galleries, context menus, backstage pages and commands outside the ribbon are retained. These counts are not unique-feature counts, completion percentages or a claim that Office 2016 includes every later Microsoft 365 feature.

## Acceptance rules

Each generated checkbox is a full command acceptance gate. A partial implementation stays unchecked even if a tested subset exists. Record verified subsets in the root [TASKS.md](../../TASKS.md) and scope reports. Certifying an entire gallery/dialog requires its options and applicable editing, undo, persistence, import/export and error behavior to pass. View-only controls need display/interaction and non-mutation evidence; native compatibility claims require independent Office evidence.

The generated files intentionally preserve existing checked rows when regenerated against the pinned source. They do not infer completion from a matching label or source-code function. Any future baseline revision needs a reviewed migration of row identities and evidence.

```powershell
node scripts/download-ribbon-catalog.mjs
node scripts/build-ribbon-checklists.mjs
node scripts/build-ribbon-checklists.mjs --check
node scripts/test-ribbon-checklists.mjs
```

The downloader verifies pinned Git blob hashes; generation verifies the source SHA-256 and emits every named source row. `--check` fails on checklist drift. Downloads need network access; subsequent generation and checking work locally. Generated checklist Markdown is excluded from Prettier to retain deterministic row identities.

## Word tab map

| Screenshot tab | English tab | Catalog section                                                                      | Key remaining command families                                                                                        |
| -------------- | ----------- | ------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------- |
| Datei          | File        | TabInfo, TabOfficeStart, TabRecent, TabSave, TabPrint, TabShare, TabPublish, TabHelp | Backstage settings, protection, inspection, native print fidelity and browser alternatives for desktop integrations   |
| Start          | Home        | TabHome                                                                              | Complete clipboard/Paste Special, font dialogs/effects, full numbering, sorting, borders/shading and style management |
| Einfügen       | Insert      | TabInsert                                                                            | Cover/blank pages, advanced tables, shapes/charts/SmartArt, headers/footers, text objects, fields and equations       |
| Entwurf        | Design      | TabWordDesign                                                                        | Themes, style sets, effects, watermarks and page borders                                                              |
| Layout         | Layout      | TabPageLayoutWord                                                                    | Sections, columns, line numbers, hyphenation and object arrangement                                                   |
| Verweise       | References  | TabReferences                                                                        | Contents, citations, captions, indexes, cross-references and full footnote/endnote creation                           |
| Sendungen      | Mailings    | TabMailings                                                                          | Envelopes, labels, recipient lists, merge fields, preview and generation                                              |
| Überprüfen     | Review      | TabReviewWord                                                                        | Native-grade proofing, translation alternatives, comments, revision authoring/acceptance and comparison               |
| Ansicht        | View        | TabView                                                                              | Read/outline/draft modes, accurate page views, window management and macros                                           |

Contextual picture, table, drawing, chart, equation, header/footer and other tabs are included farther down the Word catalog. [WORD_HOME.md](../WORD_HOME.md) records the implemented Home/View/Review subset and its limits. Adding tab labels alone does not complete a tab.
