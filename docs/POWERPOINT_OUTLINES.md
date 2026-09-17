# PowerPoint shape outlines

The browser now renders and edits solid outlines on text boxes, rectangles and ellipses. The inspector exposes color, No outline, width, transparency and eleven DrawingML dash presets. Edits use existing undo/history and IndexedDB persistence. The same SVG outline component renders the canvas, thumbnails and slideshow without changing object bounds or intercepting selection.

`drawing-colors.ts` resolves direct line properties and supported theme `lnRef` styles, including placeholder colors and per-property overrides. Shape fill and outline fill are independent. `slide-outline.ts` supplies preset stroke patterns; `model.ts` validates optional outline color, width, opacity and dash metadata. Width is measured in normalized 960-wide canvas units, like existing slide font sizes; imported/exported EMU widths use the actual source slide width. The control does not claim to display native point units for arbitrary source dimensions.

Retained exports patch only changed line fields. Source line caps, joins, compound settings, effects and relationships are retained. New text/shape exports also emit outline color, width, opacity and presets. Unsupported source line paints/custom dashes remain in retained XML, but do not have complete browser rendering. Imported picture outlines may render; editing/new-picture outline export is not covered by this delivery.

## Verification

- Fourteen cases in `src/slide-outline.test.ts` cover theme/direct resolution, rejected unsupported or invalid values, no-fill independence, retained package editing and all eleven new-export dash presets on text and shapes.
- `tests/slide-outlines.spec.ts` checks a real imported no-fill rectangle with a red border, edits all four outline fields, tests No outline/undo, saved reload and slideshow display, then downloads/reimports the actual PPTX and compares every unrelated ZIP payload. Screenshot: `test-results/pptx-outlines.png`.
- `scripts/verify-powerpoint-outline-export.ps1` opens both actual browser fixtures read-only in desktop PowerPoint, verifies color, width, transparency and the edited dash-dot preset, asserts untouched background/fill/second slide, and renders PDFs. Its SHA-256 receipt is `.local/pptx-validation/outlines-report.json`. Native checking covers the authored solid/dash-dot examples, not visual identity for all presets.

```powershell
npx.cmd vitest run src/slide-outline.test.ts
npx.cmd playwright test tests/slide-outlines.spec.ts
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-powerpoint-outline-export.ps1
```

Native automation is test-only, disables macros, restores settings and closes its own presentations. No Office installation is used by the shipped static application.

## Remaining limits

This does not complete native line rendering. Custom dash arrays, compound strokes, inset alignment, exact cap/join behavior, arrowheads, zero-width hairlines, gradient/pattern lines, arbitrary geometry and full group/placeholder inheritance remain incomplete. SVG dash spacing is an approximation of desktop rendering; native field equality is not pixel equality. Old saved imports need reopening to populate source outlines. Complete Office desktop parity remains unverified and incomplete.

Primary sources: [Microsoft DrawingML outline schema](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outline?view=openxml-3.0.1), [preset line dash values](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetlinedashvalues?view=openxml-3.0.1), [PowerPoint LineFormat properties](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.lineformat).
