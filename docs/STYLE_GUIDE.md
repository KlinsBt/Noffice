# Interface style guide

Insert > Section break also accepts flat numbered/bulleted items, preserves selected text and resumes typing after the boundary. An empty ending list paragraph has no marker or number increment; an empty following item still counts. Undo/redo and reload derive the same markers. List/table pagination is still unsupported and retains the existing continuous-view explanation; this is an application limitation, not a completed layout feature.

Body Layout and the header/footer editor share alignment, point indents and paragraph-break controls. Mixed values stay blank/indeterminate; valid changes are atomic, invalid indents show feedback, Escape restores the displayed selection value, and no-op changes add no history. Physical indent input uses measured native Single/twip rounding through ±1,584pt. Tab commits without moving focus back into the editor; Enter commits and returns to the editor. Indented stories now keep measured page geometry; full font/justification/overflow equivalence remains required.

Word paragraph Before/After controls are shared by the body and header/footer editors. Each side offers Points, Lines and Auto; mixed selections stay blank until a value is chosen. Invalid input leaves the document unchanged. The same-style spacing checkbox can show an indeterminate state. One field change is one history event, and ordinary numeric changes retain input focus. Contextual suppression and measured baseline paint are view details; saved content retains the authored spacing values. [Measured scope and remaining contexts](parity/contracts/W-TEXT-UNIFORM-LINES.md#current-executable-handoff).

Word line-spacing presets and custom Multiple, Exactly and At least options use the same controls in the body and header/footer editors. Custom story options expand inside the existing story dialog. Mixed selections remain blank; invalid input disables Apply, and applying an unchanged value adds no undo event. Cancel preserves the selection and content. One valid change is one history event. Paragraph-origin paint corrections remain derived view data and never enter stored content.

Word retains semantic Tab characters separately from view-only measured spacers. Eligible body and header/footer fields use imported custom/default/decimal stops; real Tab typing participates in ordinary history and reload. A verified single physical line keeps Home/End selection across tabs. Never serialize spacer widths or baseline paint into document content. Bar stops paint black across the paragraph's line boxes, including empty paragraphs, independently of literal Tab characters and text color. Paginated body bars follow physical fragments; print uses the corresponding disposable fragment tree and white paper. Unmeasured or leader tabs cannot silently disappear from direct PDF; unsupported wrapping, indents and stop authoring remain explicit contract gaps. Saved documents with missing story provenance must retain working backup/original recovery downloads.

Headers and footers > Page options edits the selected section's first-page flag and distances; odd/even selection is document-wide. Distances use points with0.05pt rounding. Section links selects the section, header/footer and page type; the first section cannot link to a predecessor. Apply is one document history event; Cancel leaves the file unchanged. Linking uses the previous section's corresponding content, while unlinking keeps an independent editable copy. Imported empty slots and fresh single-section document slots open in the existing story editor when their style template is available. Untouched blank Apply is a no-op; typing creates the story. Inactive first/even slots retain their text, with guidance to enable the corresponding page option. Fresh stories inherit the document Normal font with Single spacing and zero space after. Legacy documents whose edited appearance cannot be migrated safely show working backup/original downloads in a recovery view. Broader story authoring and full interaction/print acceptance remain open.

Eligible imported continuous sections and equal columns use the same text editor and existing spacing/keep controls. Insert exposes Page break and Column break, with Ctrl+Enter and Ctrl+Shift+Enter, plus a Section break selector for Next Page, Continuous, Even Page and Odd Page. A section break inserts before the selection and preserves its text, then returns the caret immediately after the boundary. Unsupported selection contexts show feedback without changing the document. The tested native commands split paragraphs and place the cursor for continued typing; Undo restores the state before the command. Pointer edits, history and reload act on semantic paragraphs across visual fragments. Page displacement and print copies are internal view details. Unsupported layouts retain the continuous-view explanation. The full font, command-context, header/footer and interaction matrix remains required. [Current scope](parity/contracts/W-TEXT-UNIFORM-LINES.md#current-executable-handoff).

Scalar Excel references retain ordinary cell/formula-bar editing. Caller-aligned blanks display numeric zero, formula-produced empty text stays empty, and an absent intersection reports `#VALUE!`; history and reload use the existing transactions. PowerPoint group-bound normalization happens during retained export and adds no new selection control. Whole-group browser authoring remains open.


Supported children inside rotated/reflected PowerPoint groups use the existing position/size controls and history. Geometry fields describe the visible child frame. Unsupported child rotation/reflection changes, text and skew produce an error before committing; do not present every rotated group as uneditable. Keep native intermediate Size-panel semantics distinct from the tested final-frame/export contract.

For eligible rotated group children, numeric Width/Height edits now retain the rotated upper-left corner, shifting the frame's X/Y in the same history event. Set dimensions before positions when requesting a specific final frame. Pointer resize and explicit model-frame commands remain distinct; their complete native gesture matrix is open. Browser fields use canvas units; the installed German pane's centimeter rounding is not silently applied to them.

Direct entry below eligible Excel tables expands the range without a dialog. Formula entry activates a totals row and restores supported labels; structured data sums exclude it. Undo restores the entered value, range and totals together. Existing neighboring worksheet data prevents expansion; unsupported table structures retain ordinary entry behavior.

Calculated-table resizing stays in Table design's existing range field. Apply changes the range and generated formulas together; Undo restores both. Preserve values already below the table and explain rejected ranges in the dialog before committing.

Calculated columns use direct in-cell/formula-bar entry without another dialog. One undo restores the whole fill. Preserve explicit exceptions, and report validation/protection failures before changing any body cell. Master formulas remain document metadata, not extra editor labels.

Direct table-heading edits use the ordinary cell/formula-bar and paste flows. During asynchronous reference repair, show an updating indicator, temporarily disable spreadsheet interaction and Export, and resume focus after completion. Keep the header/data rectangle as one undo step; error messages must explain rejection without committing a partial paste.

## Principles

The bounded Word baseline correction keeps the ordinary text selection, caret and font/spacing controls. Its profile, paint transforms and invisible struts are view details; do not expose calibration controls or store them in the document. Unmeasured font/spacing combinations retain the existing rendering path, with broader compatibility limits documented in the contract.

Leading, trailing and consecutive hard-break lines keep the ordinary Word editor caret. Measured struts are invisible, noneditable view decorations; never expose them as document content. Typing on an empty final line uses the paragraph-mark font, and Undo/Redo preserves explicit typing/paste fonts. [Bounded empty-line behavior](parity/contracts/W-TEXT-UNIFORM-LINES.md#empty-hard-break-lines).

Source page chrome must stay outside the specified text rectangle: an outline may frame the page without consuming its body width or margins, and must disappear in print. Imported paragraphs preserve spaces and line breaks while allowing trailing wrap-spaces to hang. Uniform text uses the existing editor and caret when line spacing changes; do not create separate editable line widgets. [Measured scope](parity/contracts/W-TEXT-UNIFORM-LINES.md#uniform-wrapped-line-subset).

Home > Line spacing > Line spacing options opens a focused dialog for Multiple, Exactly and At least. Label amounts as lines or points, validate before applying and preserve the paragraph selection on Cancel/Apply. A minimum value displays as At least in the selector. Keep spacing in the ordinary document Undo/Redo history. Legacy imported documents finish source-spacing hydration before accepting edits.

The existing Layout widow/orphan checkbox repaginates eligible exact-spaced paragraphs and participates in text Undo/Redo. Empty continuation lines keep the original editor caret and support direct entry, including at reduced zoom. Imported exact line spacing displays its physical value in the line-spacing selector. Eligible overflow fragments remain part of one editable paragraph across pages; preserve keyboard and pointer interaction with the same text.

Eligible mixed sections display centered white pages at their own dimensions with 24px between pages. Preserve one editable text surface and the existing ribbon. Keep horizontal overflow inside the document scroller. A continuous-view status identifies layouts that cannot yet be paginated; do not render invented page boundaries or clip overflowing text.

Imported single-section custom page sizes and margins display as custom in Page setup. Choosing a preset explicitly applies it to the whole document; Undo restores the prior custom geometry. Keep these choices in the same undo history as editing. Preserve the existing responsive surface at narrow widths; do not label a supported page rectangle as full Word pagination.

Familiar editing tools, a calm workspace, and direct labels. The landing page contains three mode choices. Recent files belong in an Open dialog. Surface the next useful action without making every desktop feature a visible button. Every enabled control must work.

## Visual system

| Token       | Value            | Purpose                             |
| ----------- | ---------------- | ----------------------------------- |
| Ink         | `#29382f`        | Main text                           |
| Primary     | `#2d493c`        | Main actions and workspace identity |
| Canvas      | `#f9faf7`        | Workspace background                |
| Subtle fill | `#f1f3ec`        | Secondary surfaces                  |
| Border      | `#e6e9e2`        | Low-contrast structure              |
| Word        | muted blue       | Document identity                   |
| Excel       | muted green      | Spreadsheet identity                |
| PowerPoint  | muted terracotta | Presentation identity               |

Use locally bundled Inter for UI text; Georgia is reserved for editorial artwork. Do not fetch fonts from a third-party host. Icons use the locally bundled Lucide Svelte package. All decorative illustrations are CSS or local vector markup.

Use 4–8 px spacing increments, 6–12 px corner radii, restrained shadows, and clearly visible focus rings. File actions must remain accessible at mobile widths.

## Editor ribbons

Excel Format as table/Table design is a contextual action beside Home. Its dialog puts name and range side by side when space permits, pairs style toggles, and right-aligns Apply/Cancel. It does not add another crowded Home ribbon group.

Table design pairs its column selector and editable heading on a second row, defaulting to the selected cell's column. Apply exposes progress and prevents duplicate submissions while the workbook's references are repaired. The fields stack only below 480 px.

The spreadsheet uses 6 px group gutters and 4 px control gaps. Whole-row/column operations share a single menu so the expanded command set stays horizontal at 1920 px. Wide-layout checks include color controls as well as buttons and fields.

Use `RibbonGroup.svelte` and `src/ribbon-layout.css` for all editor ribbons. Each tab has one toolbar containing named groups, with a caption below the controls. Groups use intrinsic minimum/content widths and share the available space up to their natural single-row width. Controls flow horizontally and wrap only when their actual available width requires it. Review and View occupy the same ribbon as the other Word tabs.

- Use 12 px control text, 13 px navigation text and 11 px group captions. Do not shrink controls or labels to fit a window.
- Keep icon buttons 30 px square, number fields at least 64 px wide, and dropdowns wide enough for their normal values. Keep units beside short visible labels; retain full accessible names.
- Use neutral blue-gray surfaces, fine separators and clear hover/focus states. Selected tabs use both an underline and a tinted background. Word, Excel and PowerPoint use blue, green and terracotta accents respectively.
- Do not prescribe fixed group widths, forced two-row wrappers or vertical Undo/Redo stacks. Expand groups into available width before wrapping their controls; wrap whole groups only when their minimum usable widths no longer fit. Keep labels beside Layout number fields to avoid tall forms. The ribbon must not create horizontal page overflow or squeeze inputs. At narrow widths, bound the ribbon height and allow vertical scrolling so the document remains usable and every control remains reachable.
- Keep action buttons such as Find and Filter visually distinct from selected navigation tabs. Do not add placeholder tabs for unimplemented commands.
- Hide ribbon navigation and controls in print output. Preserve editor selection when applying Word ribbon commands.

The layout regression in `tests/ribbon-layout.spec.ts` checks all five implemented Word tabs, contextual table controls, and the Excel/PowerPoint ribbons at desktop, tablet and phone widths. Screenshots support visual review; they do not certify desktop Office layout fidelity.

## Navigation

Transient notification text must not intercept clicks on the document beneath it. Its dismiss button remains interactive and keyboard accessible. Editor interactions such as table resize boundaries must still give the editing surface keyboard focus.

There is no platform dashboard or left rail. The landing page offers Word, Excel, and PowerPoint, with Open a file and Recent files as secondary actions. An open editor exposes app switching, a file name, accurate save status, version history, and export. The file library supports search, type filters, sorting, list/grid views, and reversible trash.

## Content and accessibility

- Give icon-only buttons accessible names and disabled states.
- Use native buttons, inputs, selects, and modal dialogs.
- Give editors descriptive region/textbox names; support keyboard paths for core interactions.
- Use explicit error messages. Never display "Saved" before the transaction commits.
- Explain format losses at import and export. Avoid Microsoft logos or claims of official affiliation.
- Honor reduced-motion preferences. Never communicate selection or failures through color alone.
- Automated accessibility checks complement, but do not replace, keyboard and screen-reader review.

## Font controls

Spreadsheet Ctrl+Shift+Enter commits a supported single-cell array formula and keeps that cell active. Show `{=formula}` in the inactive formula bar and raw `=formula` while editing. Ordinary Enter removes array intent on reentry and advances after in-cell editing. Reject unsupported multi-selection or table array entry with a clear message and preserve the previous value/selection. History restores both the formula and its entry intent.

For imported PowerPoint children with coarse source coordinates, position/size fields commit on change (Tab or blur), so typing a multi-digit value does not snap each keystroke. Show the resulting value after commit, retain explicit failure messages and avoid history entries for snapped no-ops. Ordinary fine-coordinate fields retain their existing live input behavior. Browser canvas units and rounded inspector displays remain distinct from the native centimeter fields; the complete native unit-entry interaction matrix is still open.

Separate bundled fonts from system fonts that may be unavailable. Use the shared FontPicker across all three modes. Word and spreadsheet sizes are points; the slide model uses 960×540 canvas units. Preserve requested family names in models and exports. Never imply that a listed system family is installed. Do not fetch remote font assets.

Word Layout page setup explicitly selects Whole document or Current section. Current-section controls follow the document selection; selections spanning sections disable the fields with an explanatory status. Property changes keep the body selection and form one history event, and the following text edit is separate. Full native dialog behavior and additional selection scopes remain in the compatibility contract.
