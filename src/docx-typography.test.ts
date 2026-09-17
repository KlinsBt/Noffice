import { it, expect, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { Editor } from '@tiptap/core';
import { readDocx } from './docx-import';
import { wordExtensions } from './word-extensions';
import { sanitizeHTML } from './formats';
import native from '../tests/fixtures/native-word-tab-leaders.json';

vi.mock('mammoth', async (original) => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});

it('preserves native explicit RGB on text and semantic tabs without changing source bytes or leaking reading styles', async () => {
  const bytes = readFileSync('tests/fixtures/word-tab-leaders.docx'), input = Uint8Array.from(bytes).buffer;
  const { content, messages } = await readDocx(input);
  expect(Buffer.from(input)).toEqual(bytes);
  const editor = new Editor({ extensions: wordExtensions(), content: sanitizeHTML(content.html) });
  try {
    expect(editor.state.doc.childCount).toBe(native.cases.length);
    for (const [index, sample] of native.cases.entries()) {
      editor.state.doc.child(index).forEach((node) => {
        expect(node.marks.find((mark) => mark.type.name === 'textStyle')?.attrs.color)
          .toBe(sample.color === 'FF0000' ? 'rgb(255, 0, 0)' : 'rgb(0, 0, 0)');
      });
    }
    expect(editor.getHTML()).not.toMatch(/NOFFICE|noffice-font-/);
    expect(messages.filter((message) => message.message.includes('NOFFICE'))).toEqual([]);
  } finally { editor.destroy(); }
});
