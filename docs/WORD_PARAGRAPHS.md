# Word paragraph layout controls

The Layout tab supports indentation before/after text, signed first-line indentation (negative for hanging), keep-with-next, keep-lines-together, page-break-before and widow/orphan control. Controls apply to selected paragraphs/headings. Each layout action is a separate undo step. Values persist as semantic inline CSS in the native document model.

The importer resolves supported style inheritance before reading Word's indentation and paragraph flags. The browser displays indents and exposes break constraints to its print engine. Newly generated DOCX files write the corresponding WordprocessingML properties. Retained DOCX exports patch only changed paragraph properties in source order, preserving unrelated payloads and explicit false overrides. Changing indents clears corresponding character-unit overrides using the zero values described by Microsoft's implementation notes.

## Verification

`src/docx-preserve.test.ts` covers direct indentation, explicit on/off flags, inherited controls, new-DOCX generation, retained-part comparisons and reimport. It also checks explicit refusal of the unresolved inherited-hanging-to-first-line conversion. The source run contents, headers and unrelated package parts must remain intact.

`tests/docx-fidelity.spec.ts` selects a source paragraph, changes indents and break settings, checks CSS, undoes one setting, saves/reloads, downloads DOCX and reimports it. This caught and fixed Tiptap history grouping multiple independent layout changes into one undo step.

```powershell
npm.cmd run test -- src/docx-preserve.test.ts src/docx-styles.test.ts
npm.cmd run test:e2e -- tests/docx-fidelity.spec.ts
```

## Remaining scope

The editor still uses continuous layout. CSS print constraints do not reproduce Word's full pagination, line metrics or page-break decisions. No new native Word rendering certification is asserted. Character-unit indentation rendering, list indentation interactions, mirrored/complex-script layout and the full style cascade remain incomplete. Changing an inherited hanging indent into a first-line indent is rejected until style-aware export can override it reliably. Other supported direct indentation edits work without rebuilding the document.

Primary sources: [Microsoft paragraph indentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.indentation?view=openxml-3.0.1), [Word character-unit override behavior](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/138732dc-507a-4164-af66-808c2fd9f2f9), [keep-with-next semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.keepnext?view=openxml-3.0.1).
