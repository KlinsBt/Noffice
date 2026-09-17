# Word hyperlinks

September 9 regression follow-up: immediate native keyboard selections can precede ProseMirror's selection observer. `syncWordSelection` now synchronizes in-document DOM anchor/head positions after keyboard/mouse selection events, before ribbon actions and before the link dialog takes focus, preserving forward/backward selection direction. Two new unit regressions cover capture; the unchanged Ctrl+K, formatting and View workflows pass three repetitions each. Repeated full runs exposed the initial dialog-only fix as insufficient for ribbon formatting; selection synchronization now happens before either consumer needs it.

Insert > Link and Ctrl+K now add a web/email link at the cursor or over selected text, edit an existing link from within it, change its display text, and remove its destination. Each operation has its own undo step. Unchanged selected text retains mixed formatting; replacement display text inherits the first character's marks. After insertion the caret returns to the editor and subsequent typing is outside the link.

## DOCX handling

`word-links.ts` owns selection and editor transactions. `docx-links.ts` manages hyperlink wrappers and relationship IDs independently for the body, footnotes and endnotes. The retained writer preserves source run properties and unchanged hyperlink attributes. Changing one destination creates/reuses a relationship without retargeting other links that shared the original relationship. Unrelated package payloads remain intact. New-document exports also retain mixed formatting within links.

The editor and retained comparison now both preserve leading/trailing HTML whitespace. Previously a trailing source space could disappear on parsing, preventing export from mapping a paragraph after cursor insertion. Dialog application restores focus synchronously after Svelte removes the modal, avoiding deferred focus that could discard following typing.

## Evidence

- `src/word-links.test.ts`: selection/cursor editing, mixed formatting, validation, separate undo/redo and following typing.
- `src/docx-links.test.ts`: five cases covering retained formatting/attributes, shared targets, independent footnote relationships and reference markers, newly generated rich links and added linked paragraphs.
- `tests/word-links.spec.ts`: actual keyboard selection, Ctrl+K, edit/remove, undo/redo, invalid-address rejection, insertion, following typing, saved reload, DOCX download and reimport; all unrelated ZIP payloads compared.
- `scripts/verify-word-links.ps1`: installed Microsoft Word verifies the actual browser export's three destinations, labels and all ten characters' bold/italic state, then renders a PDF. The receipt binds the result to the exported SHA-256.

```powershell
npx.cmd playwright test tests/word-links.spec.ts
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-word-links.ps1
```

The native script needs desktop Office access. It opens only the local authored export read-only, disables macros, never follows the links, restores application settings and does not quit a previously running Word instance. Local artifacts: `.local/docx-validation/links.docx`, `links.pdf`, `links-report.json`.

## Remaining acceptance

This covers web/email targets and supported simple paragraphs. Complete bookmark/file-target creation, ScreenTip editing, heading navigation, visited-link appearance, all hyperlink field forms and complex bookmark/review/object structures remain open. Existing internal hyperlink wrappers can be retained; there is no complete bookmark-target authoring UI. Full Word layout and full hyperlink-command parity are not certified. Complex paragraph structures continue to fail explicitly when unsupported edits cannot be preserved.

Primary references: [Microsoft create/edit hyperlink](https://support.microsoft.com/en-us/word/create-or-edit-a-hyperlink), [keyboard hyperlink guidance](https://support.microsoft.com/en-us/accessibility/word/use-a-screen-reader-to-insert-a-hyperlink-in-word), [Word field result ranges](https://learn.microsoft.com/en-us/office/vba/api/word.field.result). Native character assertions use exact result offsets: Word's `Characters.Item(1)` can include hidden field-code positions even when obtained through a displayed result range.
