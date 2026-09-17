# PowerPoint solid fills and backgrounds

Text boxes, rectangles and ellipses now import and display supported solid fills and transparency. Fill color, No fill and Fill transparency controls edit those properties with undo and persistence. Text-box fills are also emitted by the new-presentation writer, which previously omitted them. The canvas, thumbnails and slideshow share the same paint conversion.

## Resolution and preservation

`drawing-colors.ts` separates fill resolution from rendering. It reads direct fill properties without selecting colors from a shape's line or nested descendants. It resolves sRGB colors, system-color `lastClr` values, theme scheme colors, master color maps and layout/slide overrides. Supported fill/background style references resolve `phClr` against the theme fill matrix. Missing direct fills can inherit from mapped layout/master placeholders. Backgrounds resolve from slide, layout and master; absent backgrounds default to white.

Supported transforms are tint, shade, luminance/saturation value/modulation/offset and alpha value/modulation/offset, applied in source order. Tint/shade operate in linear sRGB; luminance/saturation operate in HSL. Color-transform chains are bounded at 64 entries. Unresolved colors and unsupported fill types do not borrow arbitrary descendant colors.

Models store optional `fillOpacity` and `backgroundOpacity`; older values default to opaque where needed. Retained PPTX exports patch only changed paints and preserve the source theme/background parts when a shape fill is edited. RGB edits become direct colors; unchanged source scheme references remain. Hex case alone does not cause a paint rewrite. Reopen older imported decks to load resolved source appearance.

## Verification

- `scripts/generate-paint-oracle.mjs` creates 24 authored RGB/transform cases and a local PPTX; `scripts/verify-powerpoint-paints.ps1` records native RGB and opacity results, bound to the input JSON and presentation hashes.
- `src/drawing-colors.test.ts` verifies all 24 native cases, receipt binding, color maps, style references, inheritance, unsupported inputs and no-fill precedence. Color tolerance is at most one 8-bit channel step because native Office quantizes transforms; opacity tolerance is 0.000001.
- `src/pptx-paints.test.ts` verifies actual package import, text-fill/alpha editing, new-deck export and unrelated ZIP payload retention.
- `tests/slide-paints.spec.ts` checks source colors, opacity editing, No fill/undo, saved reload, slideshow rendering, actual export/reimport and every unrelated source payload. Screenshot inspected.
- `scripts/verify-powerpoint-paint-export.ps1` opens the actual source and browser-edited exports read-only in native PowerPoint. It asserts the inherited background, edited text fill/opacity, untouched no-fill/theme objects and second slide, then renders both PDFs. The hash-bound receipt is `.local/pptx-validation/paints-report.json`.

```powershell
node scripts/generate-paint-oracle.mjs
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-powerpoint-paints.ps1
npx.cmd vitest run src/drawing-colors.test.ts src/pptx-paints.test.ts
npx.cmd playwright test tests/slide-paints.spec.ts
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-powerpoint-paint-export.ps1
```

Native automation is test-only; it disables macros, restores settings, closes its own presentations and never overwrites the source files. It is not part of the static app.

## Remaining limits

Gradient, pattern, picture and group fills, complete theme overrides, all DrawingML color kinds/transforms, theme editing, effects, mixed-run text colors and full placeholder/style inheritance remain incomplete. Supported solid outlines are documented in [POWERPOINT_OUTLINES.md](POWERPOINT_OUTLINES.md). System colors use their saved fallback rather than the current operating-system palette. Background transparency has model/export support but no dedicated editing control or complete compositing certification. Unsupported paints remain in the retained package; that is not full browser rendering support. This delivery does not certify complete PowerPoint or Office parity.

Primary references: [Microsoft slide masters and color maps](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-slide-masters), [DrawingML format scheme](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.formatscheme?view=openxml-3.0.1), [PresentationML background](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.background?view=openxml-3.0.1), [DrawingML primer and color transforms](https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%203%20-%20Primer.pdf).
