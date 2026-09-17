# PowerPoint object layering

Arrange > Order objects provides Bring to front, Bring forward, Send backward and Send to back for single or multiple selections. Selected objects retain their relative order. Each command is one undo operation; commands that cannot change the order create no history entry. The inspector's existing Bring to front action uses the same implementation.

## Model and retained export

`slide-stack.ts` owns the immutable command logic. Slides carry an optional `stackOrder`; imported elements carry `sourceStackKey`. Import reads interleaved text/shapes/pictures in XML order instead of placing all pictures after all shapes. Direct source drawing nodes are indexed by `pptx-stack.ts`, including charts and other objects the browser does not render. These remain actual layers when a supported object moves forward, backward, to the front or to the back.

The retained writer matches edited objects by source identity, then moves the existing shape-tree nodes. It preserves unrelated package payloads, object identities and relationships, including chart and media parts. Added text/shapes/pictures can be placed between retained objects. Group descendants retain their parent group's layer identity; an individual group child cannot be reordered using these commands. Invalid/missing/duplicate source layer metadata fails explicitly.

Reopen older imported presentations to load complete stacking metadata before using the commands. Their supported ordinary text/geometry edits retain the existing source order. New presentations use their ordered element array for export.

## Evidence

- `tests/fixtures/stack-oracle-cases.json`: every nonempty selection of five objects, in normal and reverse selection order, across all four commands (248 cases).
- `scripts/verify-powerpoint-stack.ps1`: executes those inputs in installed PowerPoint and records the input SHA-256 plus native order in `native-powerpoint-stack.json`.
- `src/slide-stack.test.ts`: 248 native comparisons, receipt binding and selection/identity/group/legacy checks.
- `src/pptx-stack.test.ts`: five import/export cases for mixed image/text order, opaque chart layers, editing after reordering, inserted pictures, unchanged source node/package payloads, invalid metadata and grouped descendants.
- `tests/slide-stack.spec.ts`: actual single/multiple selection, all commands, undo/redo, saved reload, inserted text, PPTX download/reimport and every unrelated ZIP payload comparison.
- `scripts/verify-powerpoint-stack-export.ps1`: opens the actual browser export read-only in native PowerPoint, asserts all five layers including the chart, verifies the untouched second slide and renders a PDF. `.local/pptx-validation/stack-report.json` binds the receipt to the file hash.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-powerpoint-stack.ps1
npx.cmd vitest run src/slide-stack.test.ts src/pptx-stack.test.ts
npx.cmd playwright test tests/slide-stack.spec.ts
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify-powerpoint-stack-export.ps1
```

Native scripts use temporary/read-only presentations, disable macros, restore their settings and close their own documents. Test artifacts and Office automation are not deployed with the browser app.

## Remaining scope

This is not complete PowerPoint parity. Group editing/ungrouping, complete selection-pane behavior, animation-aware deletion/duplication, slide master editing and rendering/editing of retained charts, tables, connectors, SmartArt and media remain open. A layer can be retained without being rendered in the browser; moving across such a layer may only become visible in native PowerPoint. The native receipt verifies selected stacking behavior, not all visual or playback fidelity.

Primary references: [Microsoft ShapeTree ordering rules](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.shapetree?view=openxml-3.0.1), [ShapeRange.ZOrder](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shaperange.zorder), [object arrangement guidance](https://support.microsoft.com/en-us/office/graphics-visuals/align-or-arrange-objects).
