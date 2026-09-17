import { Extension, generateHTML, getSchema, type JSONContent } from '@tiptap/core';
import { DOMParser as EditorDOMParser } from '@tiptap/pm/model';
import { EditorState } from '@tiptap/pm/state';
import { fixTables } from '@tiptap/pm/tables';
import StarterKit from '@tiptap/starter-kit';
import TextAlign from '@tiptap/extension-text-align';
import { TextStyleKit } from '@tiptap/extension-text-style';
import Highlight from '@tiptap/extension-highlight';
import Image from '@tiptap/extension-image';
import { TableKit } from '@tiptap/extension-table';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import { WordScriptFormatting } from './word-script';
import { ParagraphLayout } from './paragraph-layout';
import { WordParagraphSpacingView } from './word-paragraph-spacing-view';
import { WordParagraphSpacingCommands } from './word-paragraph-spacing-commands';
import { WordLineSpacingCommands } from './word-line-spacing-commands';
import { WordParagraphLayoutCommands } from './word-paragraph-layout-commands';
import { WordFontInput } from './word-font-input';
import { WordParagraphInput } from './word-paragraph-input';
import { WordBoundaryDelete } from './word-boundary-delete';
import { WordFontCommands } from './word-font-commands';
import { WordTab } from './word-tab';
import { WordHyphen } from './word-hyphen';
import { WordTabView } from './word-tab-view';
import { WordPageBreak, WordColumnBreak } from './word-page-break';
import { WordFormattingMarks } from './word-marks';
import { WordEditRun } from './word-edit-run';
import { WordKerning, WordTextStyle } from './word-kerning';
import { WordFontFeatures } from './word-font-feature-extension';
import { WordListSectionView } from './word-list-section-view';
import type { WordNumbering } from './word-list-layout';

/** References describe the retained package, never a URL or an executable instruction. */
const SourceParagraph = Extension.create({
  name: 'sourceParagraph',
  addKeyboardShortcuts() {
    const select = (position: number | { from: number; to: number }) => {
      const result = this.editor.chain().setTextSelection(position).scrollIntoView().run();
      // Refocusing after a dialog can leave Chromium's caret at the start
      // while the model is already at the requested end. An equal model
      // selection alone does not rewrite that DOM selection.
      if (result) this.editor.view.focus();
      return result;
    };
    return {
      'Mod-Home': () => select(1),
      'Mod-End': () => select(this.editor.state.doc.content.size),
      'Mod-Shift-Home': () => select({ from: this.editor.state.selection.anchor, to: 1 }),
      'Mod-Shift-End': () =>
        select({
          from: this.editor.state.selection.anchor,
          to: this.editor.state.doc.content.size,
        }),
    };
  },
  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          sourceParagraph: {
            default: null,
            parseHTML: (el) => el.getAttribute('data-source-paragraph'),
            renderHTML: (attrs) =>
              attrs.sourceParagraph ? { 'data-source-paragraph': attrs.sourceParagraph } : {},
          },
          paragraphMarkSource: {
            default: null,
            parseHTML: (el) => el.getAttribute('data-word-paragraph-mark-source'),
            renderHTML: (attrs) => attrs.paragraphMarkSource == null ? {} :
              { 'data-word-paragraph-mark-source': attrs.paragraphMarkSource },
          },
          direction: {
            default: null,
            parseHTML: (el) => (['rtl', 'ltr'].includes(el.dir) ? el.dir : null),
            renderHTML: (attrs) => (attrs.direction ? { dir: attrs.direction } : {}),
          },
        },
      },
    ];
  },
});

/** The exporter must compare the same normalized schema the user actually edits. */
export function wordExtensions(options: { retainedEditRuns?: boolean; tabPageLeft?: () => number | null; numbering?: () => WordNumbering | undefined } = {}) {
  return [
    StarterKit.configure({
      link: { openOnClick: false, autolink: true, protocols: ['https', 'http', 'mailto'] },
    }),
    TextAlign.configure({ types: ['heading', 'paragraph'] }),
    TextStyleKit.configure({ textStyle: false }),
    WordTextStyle,
    WordKerning,
    WordFontFeatures,
    ParagraphLayout,
    WordParagraphSpacingView,
    WordParagraphSpacingCommands,
    WordLineSpacingCommands,
    WordParagraphLayoutCommands,
    WordFontInput,
    WordParagraphInput,
    WordBoundaryDelete,
    WordFontCommands,
    WordEditRun.configure({ retained: !!options.retainedEditRuns }),
    WordTab,
    WordHyphen,
    WordTabView.configure({ pageLeft: options.tabPageLeft || (() => null) }),
    WordPageBreak,
    WordColumnBreak,
    WordFormattingMarks,
    WordListSectionView.configure({ numbering: options.numbering || (() => undefined) }),
    SourceParagraph,
    Subscript,
    Superscript,
    WordScriptFormatting,
    Highlight.configure({ multicolor: true }),
    Image.configure({ allowBase64: true }),
    TableKit.configure({ table: { resizable: true } }),
  ];
}

export function wordJSON(html: string): JSONContent {
  const extensions = wordExtensions();
  const schema = getSchema(extensions);
  const parse = (text: string) => {
    const root = document.createElement('div');
    root.innerHTML = text;
    return EditorDOMParser.fromSchema(schema).parse(root, { preserveWhitespace: true });
  };
  const state = EditorState.create({ doc: parse(html) });
  // The table editor performs this repair on its first transaction. Apply the same
  // normalization to both sides of the preservation comparison.
  const normalized = (fixTables(state)?.doc || state.doc).toJSON();
  if (normalized.content?.at(-1)?.type !== 'paragraph')
    normalized.content.push(schema.nodes.paragraph.create().toJSON());
  // onUpdate persists getHTML(), including serializer normalization around inline breaks.
  return parse(generateHTML(normalized, extensions)).toJSON();
}
