import { expect, it } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { wordStoryParagraphPaint } from './word-story-paint';

it('slices wrapping story paint without losing spaces, marks or source-local baselines', () => {
  const editor = new Editor({
    extensions: wordExtensions(),
    content: '<p><strong>one two </strong><em>three four</em></p>',
    parseOptions: { preserveWhitespace: true },
  });
  try {
    const before = editor.getJSON();
    const strong = editor.view.dom.querySelector('strong')!;
    const em = editor.view.dom.querySelector('em')!;
    strong.dataset.wordBaselinePaint = em.dataset.wordBaselinePaint = 'true';
    strong.dataset.wordNativeBaseline = '12.8';
    em.dataset.wordNativeBaseline = '28.8';
    strong.getClientRects = () => [{ top: 0 }, { top: 16 }] as unknown as DOMRectList;
    const clones = wordStoryParagraphPaint(editor.view, 0, [
      { from: 0, to: 4, top: 0, height: 16 },
      { from: 4, to: 14, top: 16, height: 16 },
      { from: 14, to: 18, top: 32, height: 16 },
    ]);
    expect(clones.map((c) => c.element.textContent)).toEqual(['one ', 'two three ', 'four']);
    expect(clones.map((c) => c.top)).toEqual([0, 16, 32]);
    expect(clones[1].element.querySelector('strong')!.textContent).toBe('two ');
    expect(Number(clones[2].element.querySelector('em')!.dataset.wordNativeBaseline)).toBeCloseTo(
      12.8,
    );
    expect(editor.getJSON()).toEqual(before);
    expect(em.dataset.wordNativeBaseline).toBe('28.8');
  } finally {
    editor.destroy();
  }
});
