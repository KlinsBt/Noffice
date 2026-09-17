# PowerPoint preservation and editing

Supported [solid outline edits](POWERPOINT_OUTLINES.md) patch only changed line color/alpha, width or preset dash fields, preserving unrelated line attributes and package parts. Native PowerPoint verifies both actual source and edited browser files. Custom/compound strokes and complete browser line rendering remain open.

Imported PPTX edits now patch the retained source package. Supported existing text, direct formatting, geometry, simple fills, slide backgrounds and speaker notes no longer rebuild the presentation through the general writer. Unchanged exports still return original bytes. New presentations use PptxGenJS.

## Implemented scope

- Supported solid text/shape fills and opacity import through direct/theme/mapped inheritance and export into new or retained presentations. Source backgrounds follow slide/layout/master precedence instead of using the app canvas color. Unchanged theme references remain. [Paint evidence and remaining limits](POWERPOINT_FILLS.md).

- All four object-layering commands preserve selected objects' relative order, imported shape identities and unrelated source parts. Interleaved text/image import follows source order; charts and other opaque objects retain their layer positions. Actual browser exports pass native PowerPoint stacking checks. See [POWERPOINT_STACKING.md](POWERPOINT_STACKING.md) for group and rendering limits.

- Multiple represented objects can be selected, aligned, distributed and moved together while retaining source identities and order. The [arrangement workflow](POWERPOINT_ARRANGE.md) has 80 native PowerPoint comparisons and checks an actual browser export in desktop PowerPoint. Unsupported groups and drawing types remain outside this coverage.

- Existing text changes retain unaffected run properties and characters, including mixed font/emphasis/color runs. The editor exposes one style per object; it does not yet render or select individual rich runs. Explicit object formatting changes apply to all runs.
- Position, size and rotation changes use the original slide dimensions. Rotation appears on the canvas and slideshow and survives source-preserving and new-deck export. Resize geometry keeps the opposite corner fixed at the tested angles, including 4:3 slides. Source flip and other untouched properties remain in the package; grouped geometry changes are rejected. Older native models without rotation metadata retain the source rotation on export.
- Added text boxes, rectangles, ellipses and local PNG/JPEG/GIF/WebP images receive unique object/relationship identifiers. Images use new media parts and content types.
- Existing notes are patched at their body placeholder. First notes can create a notes slide, including on a deck without a notes master. The minimal new notes page is not full notes-print layout support.
- Slide placeholders resolve geometry from the matching layout index, then the master placeholder type. Direct geometry takes precedence. This is geometry inheritance, not complete theme, text-style or master rendering.
- Imported proportions are retained in the model, canvas, thumbnails and slideshow. Coordinates remain normalized to 960 by 540 for saved-model compatibility; display height and pointer movement account for the source ratio. New-deck generation also respects saved proportions.
- Imported slides can be reordered by moving their existing slide-ID entries; slide parts and relationships keep their identities. Editing a reordered slide still targets its original part. Decks with sections are explicitly rejected for reorder until section-aware editing is implemented.
- Unknown package parts, including charts, masters, media and timing, remain untouched unless their supported mapped content is edited.

Imported slide insertion/deletion/duplication, existing object deletion/duplication, source image replacement, grouped geometry and text changes involving fields or soft breaks fail explicitly. The source original and native backup remain available. Reopen older imports to obtain current placeholder positions, aspect and stacking metadata.

## Evidence and reproduction

`src/pptx-edit.test.ts` covers existing rich runs, direct formatting, geometry, paragraph changes, existing/first notes, images, structural rejection, placeholder precedence and custom proportions. It compares untouched ZIP payloads and reimports output. `src/pptx-preserve.test.ts` covers the existing additive path.

`tests/pptx-fidelity.spec.ts` imports a 4:3 fixture with rich text and a real chart. Browser workflows edit text, position, fonts and notes, save/reload, inspect canvas/slideshow proportions, and download the actual exports. A second workflow inserts an image and first notes into a deck without notes parts. Outputs are in `.local/pptx-validation/`.

`corpus/pptx-edits.spec.ts` edits existing text in pinned external `WithMaster.pptx`, `SampleShow.pptx` and `bar-chart.pptx` inputs. Every unrelated part must remain byte-identical, and the edited text must appear on reimport. These complement the 20 existing additive PPTX corpus workflows.

```powershell
npm.cmd run verify
npm.cmd run corpus:test -- --grep pptx
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-powerpoint-edits.ps1
```

The native helper opens the two actual browser exports read-only in installed Microsoft PowerPoint 16.0 with macros disabled. It checks slide count, 720 by 540 point dimensions, edited text/position/rotation/notes, inserted image and retained chart, then renders PDFs. Its SHA-256-bound receipt is `.local/pptx-validation/desktop-report.json`. It restores its settings and closes its own presentations. This is selected native compatibility evidence, not a full visual comparison or desktop parity certification.

## Remaining desktop scope

Full rich text and bullets, themes/master editing, complete placeholder/style inheritance, groups, text/shape flip editing, connectors, arbitrary shapes, image crop, chart/table/SmartArt editing, video/audio playback, animation/transitions, presenter view, relationship-aware slide structure editing, full typography and print fidelity remain open. Preserved objects are not automatically editable or rendered in the browser.

Rotation validation extends the actual browser export with a 30-degree title checked by native PowerPoint. `src/slide-geometry.test.ts` verifies fixed-corner resizing at 0, 30, 90 and 270 degrees. `src/pptx-edit.test.ts` includes new-deck rotation and older-model retention. Reopen older imports to display source rotation in the editor.

## Imported slide order

`src/pptx-edit.test.ts` checks reordering plus an edit to the correct source slide, unchanged unrelated payloads, duplicate-identity rejection and the section guard. `tests/pptx-fidelity.spec.ts` moves an imported slide, saves/reloads, downloads and reimports the reordered deck. Only `ppt/presentation.xml` changes when the sole operation is reordering.

`scripts/verify-powerpoint-order.ps1` opens the actual `.local/pptx-validation/reordered.pptx` read-only with macros disabled. It confirms the retained chart is on slide 1 and the original title is on slide 2, then renders a PDF. `.local/pptx-validation/order-report.json` binds those assertions to the exported SHA-256. The script restores its settings and closes only its own presentation. This establishes selected native reordering compatibility, not full section/custom-show/presentation parity.

Primary source: [Microsoft's slide-reordering guidance](https://learn.microsoft.com/en-us/office/open-xml/presentation/how-to-move-a-slide-to-a-new-position-in-a-presentation).

## Image flipping

Image objects import horizontal/vertical flip flags and expose Flip horizontally / Flip vertically. Canvas images mirror within their rotated frame; thumbnails and slideshow use the same center-based geometry. Resize handles stay on the frame. Supported existing pictures patch only their source transform, preserving media and unrelated parts. Added pictures and new decks also export the flags. Older native models without flip metadata retain the source flags when another property is edited; reopen the original to display those flags.

Two unit workflows cover imported flips, legacy-model retention, new decks and image additions. The browser workflow uses an authored four-color image, checks canvas and slideshow transforms, undo/redo, saved reload and exported/reimported flags, and compares every unrelated package payload. Undo now preserves valid object selection and focuses the editor so shortcuts keep working when an inspector control disappears.

`scripts/verify-powerpoint-flips.ps1` opens the actual `.local/pptx-validation/flipped.pptx` read-only with macros disabled. Native PowerPoint 16.0 confirms horizontal flip off, vertical flip on and 30-degree rotation, then renders a PDF. The hash-bound receipt is `flips-report.json` in that directory. This verifies selected image transforms, not grouped geometry, crop rendering, text/shape flipping or full visual parity.

Primary reference: [Microsoft DrawingML Transform2D](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.transform2d?view=openxml-3.0.1).
