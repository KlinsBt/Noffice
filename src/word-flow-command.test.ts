import { expect, it, vi } from 'vitest';
import { Editor } from '@tiptap/core';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import native from '../tests/fixtures/native-word-ui-flow.json';
import { wordExtensions } from './word-extensions';
import { readDocx } from './docx-import';
import { paragraphGraphemes } from './word-line-measurements';
import { WordEditorSections, updateWordSectionSource } from './word-editor-sections';

vi.mock('mammoth', async (original) => {
  const actual = await original<typeof import('mammoth')>();
  return {
    ...actual,
    convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
      actual.convertToHtml(
        { buffer: Buffer.from(input.arrayBuffer) },
        options as Parameters<typeof actual.convertToHtml>[1],
      ),
  };
});

for (const [name, reference] of Object.entries(native.cases))
  it(`matches native keyboard paragraph/selection/history semantics: ${name}`, async () => {
    const bytes = readFileSync(`tests/fixtures/${reference.inputFile}`);
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(reference.inputSha256);
    const { content } = await readDocx(Uint8Array.from(bytes).buffer);
    const editor = new Editor({
      extensions: [...wordExtensions(), WordEditorSections],
      content: content.html,
    });
    const snapshot = () => {
      let text = '',
        offset = 0;
      editor.state.doc.forEach((p, from) => {
        text += paragraphGraphemes(p)!.text + '\r';
        if (editor.state.selection.from >= from + 1)
          offset =
            editor.state.selection.from - 1 - editor.state.doc.content.cut(0, from).childCount;
      });
      return { text, offset };
    };
    try {
      updateWordSectionSource(editor, content.docxStructure);
      const at = reference.steps.initial.text.search(/[\f\u000e]/);
      editor.commands.setTextSelection({ from: at + 1, to: at + 2 });
      editor.commands.deleteSelection();
      expect(snapshot().text).toBe(reference.steps.deleted.text);
      editor.commands.setTextSelection({
        from: reference.steps.beforeCommand.selection.from + 1,
        to: reference.steps.beforeCommand.selection.to + 1,
      });
      const deleted = editor.getJSON();
      expect(
        editor.commands.insertWordFlowBreak(
          reference.steps.initial.text[at] === '\f' ? 'page' : 'column',
        ),
      ).toBe(true);
      expect(snapshot()).toEqual({
        text: reference.steps.inserted.text,
        offset: reference.steps.inserted.selection.from,
      });
      const inserted = editor.getJSON();
      editor.commands.undo();
      expect(editor.getJSON()).toEqual(deleted);
      editor.commands.redo();
      expect(editor.getJSON()).toEqual(inserted);
      expect(snapshot().text).toBe(reference.steps.redo.text);
    } finally {
      editor.destroy();
    }
  });
