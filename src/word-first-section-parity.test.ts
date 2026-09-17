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
import { planWordFragments } from './word-fragment-plan';
import reference from '../tests/fixtures/word-first-section-parity/reference.json';

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
  it(`matches native first-section numbering and header slots: ${sample.name}`, async () => {
    const bytes = readFileSync(sample.source);
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(sample.sourceHash);
    const { content } = await readDocx(Uint8Array.from(bytes).buffer);
    const editor = new Editor({
      extensions: [...wordExtensions(), WordEditorSections],
      content: content.html,
    });
    try {
      updateWordSectionSource(editor, content.docxStructure);
      const stories = {
        ...content.stories!,
        parts: content.stories!.parts.map((part) => ({ ...part, height: 16 })),
      };
      const input = sectionSurfaceInput(
        editor.state.doc,
        wordSectionMap(editor.state),
        resolveWordSections(content),
        stories,
      )!;
      // Each section has one short unwrapped paragraph. Height cannot change
      // these native parity decisions; text layout is covered separately.
      const plan = planWordFragments(
        input,
        input.blocks.map((block) => ({
          height: 24,
          before: 0,
          after: 0,
          keepLines: false,
          widowControl: false,
          lines: [{ from: 0, to: block.to - block.from - 2, top: 0, height: 24 }],
        })),
      )!;
      expect(plan.pages).toHaveLength(sample.native.pages);
      expect(plan.fragments.map((f) => f.page + 1)).toEqual(
        sample.native.sections.map((s) => s.physical),
      );
      expect(
        plan.pages.map((page) => {
          const path = page.stories?.find((story) => story.kind === 'header')?.path;
          if (!path) return '';
          const div = document.createElement('div');
          div.innerHTML = content.stories!.parts.find((part) => part.path === path)!.html;
          return div.textContent;
        }),
      ).toEqual(sample.headers);
    } finally {
      editor.destroy();
    }
  });
