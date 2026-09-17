# PowerPoint selection and arrangement

The Home ribbon offers six alignment commands, horizontal/vertical distribution, a slide/selection reference choice, and a selection pane. Ctrl-click, Shift-click and the pane's toggle buttons select multiple represented objects. Ctrl+A selects the current slide's editable objects when focus is outside a text field. Escape clears selection. Changing slides clears selection.

Dragging a selected object moves the selection together. Arrow keys move by one model unit; Shift+Arrow moves by ten. Objects may move beyond the slide edge. These keyboard step sizes are Noffice's model units, not certified native point increments. A drag or arrangement commits one undo step and uses the existing revision-checked local storage. Multi-selection does not create a persistent PowerPoint group.

## Geometry and native evidence

`src/slide-arrange.ts` keeps object identity, order, size, rotation, text and source metadata intact. It accounts for the source aspect ratio despite the saved model's normalized 960 by 540 coordinates. Alignment uses the rotated bounding rectangle; center/middle use the combined selection bounds or slide center. Distribution follows the measured native behavior: center ordering, equal gaps, and quadrant-snapped dimensions. Relative-to-slide distribution includes equal outer gaps. Overlapping inputs can produce negative gaps; these are retained rather than clamped.

Installed Microsoft PowerPoint 16.0 evaluated 80 authored cases on September 9, 2026: all eight commands relative to the selection and slide, with unequal sizes, overlap, 16:9 and 4:3 proportions, rotations of 30/90/270 degrees, and 44/45/46/135/225/315-degree boundaries. `tests/fixtures/native-powerpoint-arrange.json` records actual positions and binds the exact authored input bytes by SHA-256. The 82 tests in `src/slide-arrange.test.ts` compare every output within 0.005 native points, validate the input binding, and check immutable movement, identity retention and invalid selection handling. This is measured behavior for these cases, not an assertion about every shape type or PowerPoint version.

`tests/slide-arrange.spec.ts` exercises real Ctrl/Shift selections, pane keyboard toggles, alignment, distribution, group drag, keyboard movement, undo/redo, slide changes, narrow layout, saved reload and a real PPTX download/reimport. The exported 4:3 slide keeps all unrelated ZIP payloads byte-identical. The output also opens read-only in desktop PowerPoint; native positions, dimensions, rotation, text, order and slide count are asserted before PDF rendering. The receipt binds the export and oracle hashes in `.local/pptx-validation/arrange-report.json`.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-powerpoint-arrange.ps1
npx.cmd vitest run src/slide-arrange.test.ts
npm.cmd run build
npx.cmd playwright test tests/slide-arrange.spec.ts
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-powerpoint-arrange-export.ps1
```

The native scripts create a temporary presentation or open the browser-generated test export. They restore application settings and close their own presentations. They do not open user documents and do not ship in the static application.

## Remaining limits

- Selection covers editable model objects, not preserved charts, groups, masters or other unsupported drawing types. The selection pane supports selecting objects; renaming, hiding, locking and reordering remain unfinished.
- Distribution within a selection requires at least three objects. A single object aligns to the slide; the reference choice becomes available with multiple objects.
- Group resize/rotation, persistent grouping, marquee selection, snap guides and native keyboard-distance parity remain unfinished.
- Existing imported-object deletion and reordering still cannot export. Keyboard deletion now rejects imported selections before changing the document. Existing exporter guards remain in force for unsupported structures.
- Geometry comparisons and selected native export checks do not certify rich text rendering, fonts, complete slide layout or full desktop parity.

## Primary references

- [Microsoft: Align or arrange objects](https://support.microsoft.com/en-us/office/graphics-visuals/align-or-arrange-objects) describes multiple selection and alignment/distribution choices.
- [PowerPoint ShapeRange.Align](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.align) distinguishes selection and slide alignment.
- [PowerPoint ShapeRange.Distribute](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.distribute) distinguishes distribution across the existing space and the full slide. Rotation and overlap details above come from the recorded native experiments.
