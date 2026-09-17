# Word Home, View and Review commands

The Word editor now exposes additional working commands from the supplied Home ribbon, plus functional View and Review subsets. The complete command acceptance inventory is [ribbon/WORD_COMMANDS.md](ribbon/WORD_COMMANDS.md); full tab functionality remains incomplete.

## Home

- Change case: uppercase, lowercase, sentence case, capitalize each word and toggle case. Shift+F3 cycles lower/upper/title. Selected text runs keep their marks, source paragraph identities and inline objects; expanding mappings such as `straße` to `STRASSE` keep the transformed selection. Unicode context across mark boundaries, language-specific case conventions and sentence abbreviations need broader native verification.
- Grow/shrink font uses the supported 8–160 point step scale. Full desktop font-size range, every intermediate step and all font-dialog effects remain open.
- Format painter copies the first source text run's supported character marks and applies them to the next completed selection. Ctrl+Alt+C/V copy/apply repeatedly; Escape cancels painting. It replaces supported character formatting while retaining target links and paragraph identities. Paragraph/graphic painting, double-click lock, inherited style resolution and mixed-source format semantics remain open.
- Increase/decrease indent changes supported paragraph starts in 36-point increments; Ctrl+M / Ctrl+Shift+M use the same commands. Lists use nest/lift operations. Existing import/export structural restrictions still apply to complex imported lists and character-unit indentation.
- Show formatting marks (Ctrl+Shift+8) decorates spaces, paragraph ends, soft breaks and explicit tabs. Marks are editor decorations, not document characters, and are excluded from saved HTML and printing. Broader hidden/control-character display remains open.
- Select all selects the document using the editor's existing selection command.

## View and Review

View controls toggle the ruler and a heading navigation strip, set zoom from 50–150 percent and reset to 100 percent. Heading navigation uses current editor positions. These are session view preferences; exact page layout, a navigable ruler, outline editing and multi-window management remain incomplete.

Review exposes document/selection word counts, character counts and a browser spellcheck toggle. Words are counted by whitespace and character counts use JavaScript string length; these are not a complete language-aware Windows Word counting algorithm. Browser spellcheck is provided by the browser; no grammar/thesaurus, translation, comments or track-changes capability is claimed by this tab.

## Evidence

Eleven unit cases in `src/word-commands.test.ts` verify the five case modes, expansion with marks, format transfer without copying source identities/links, indentation and list nesting, non-mutating decorations and size boundaries. Two workflows in `tests/word-ribbon.spec.ts` exercise real browser selections, imported text formatting, font stepping, case changes, format painting, shortcut/undo indentation, saved reload, unchanged unrelated DOCX parts, downloaded reimport, view controls, heading navigation and selection counts.

The imported workflow verifies package preservation for its supported edits; this is not independent native Word layout or full Home-tab certification. The command catalogs remain unchecked until their full command scope has evidence.

Primary references: [Microsoft Format Painter](https://support.microsoft.com/en-us/word/use-the-format-painter), [change case](https://support.microsoft.com/en-us/Word/change-the-capitalization-or-case-of-text), [formatting marks](https://support.microsoft.com/en-us/word/show-or-hide-tab-marks-in-word) and [Word ribbon shortcuts](https://support.microsoft.com/en-us/accessibility/word/keyboard-shortcuts-in-word).
