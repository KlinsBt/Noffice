import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { Editor } from '@tiptap/core';
import { readDocx } from './docx-import';
import { wordExtensions } from './word-extensions';
import {
  WordEditorSections,
  updateWordSectionSource,
  wordSectionMap,
} from './word-editor-sections';
import { sectionSurfaceInput } from './word-section-surfaces';
import { resolveWordSections } from './word-section-layout';
import { planWordFragments, type ParagraphFlow } from './word-fragment-plan';
import reference from '../tests/fixtures/word-empty-section-end/reference.json';

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
for (const sample of reference.rows)
  it(`matches native empty section flow and typed recovery: ${sample.name}`, async () => {
    const source = readFileSync(sample.source);
    expect(createHash('sha256').update(source).digest('hex')).toBe(sample.sourceHash);
    const { content } = await readDocx(Uint8Array.from(source).buffer);
    const editor = new Editor({
      extensions: [...wordExtensions(), WordEditorSections],
      content: content.html,
    });
    try {
      updateWordSectionSource(editor, content.docxStructure);
      const check = (native: typeof sample.sourceSnapshot) => {
        const input = sectionSurfaceInput(
          editor.state.doc,
          wordSectionMap(editor.state),
          resolveWordSections(content),
        )!;
        const metrics: ParagraphFlow[] = [];
        editor.state.doc.forEach((node) => {
          const height = (parseFloat(node.attrs.paragraphLineHeight) * 4) / 3;
          metrics.push({
            height,
            before: (parseFloat(node.attrs.spaceBefore || '0') * 4) / 3,
            after: (parseFloat(node.attrs.spaceAfter || '0') * 4) / 3,
            lines: [{ from: 0, to: node.content.size, top: 0, height }],
            keepLines: false,
            widowControl: false,
          });
        });
        const plan = planWordFragments(input, metrics)!;
        expect(plan.pages).toHaveLength(native.pages);
        expect(plan.blocks).toHaveLength(native.paragraphs.length);
        plan.fragments.forEach((fragment, index) =>
          expect((fragment.top * 3) / 4).toBeCloseTo(native.nonbreakingParagraphs[index].firstY, 3),
        );
        return plan;
      };
      check(sample.sourceSnapshot);
      const before = editor.getJSON();
      let from = 0;
      for (let i = 0; i < editor.state.doc.childCount - 2; i++)
        from += editor.state.doc.child(i).nodeSize;
      editor.commands.setTextSelection(from + 1);
      editor.commands.insertContent('X');
      check(sample.typedSnapshot);
      editor.commands.undo();
      expect(editor.getJSON()).toEqual(before);
      check(sample.sourceSnapshot);
      editor.commands.redo();
      check(sample.typedSnapshot);
    } finally {
      editor.destroy();
    }
  });
