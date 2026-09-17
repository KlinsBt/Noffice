import { expect, it } from 'vitest';
import { Editor, Extension } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { wordExtensions } from './word-extensions';
import { wordFragmentPrint, wordPrintGlyphBaseline } from './word-fragment-print';
import type { SurfaceInput } from './word-section-surfaces';
import type { WordFragmentPlan } from './word-fragment-plan';

it('anchors wrapped PDF glyphs to their precise baseline plus the measured line displacement', () => {
  const p = document.createElement('p'), run = document.createElement('span'); p.append(run);
  p.style.width = '300px'; p.dataset.wordFragmentTop = '48'; p.dataset.wordFragmentSourceTop = '320';
  run.dataset.wordNativeBaseline = '345.6';
  p.getBoundingClientRect = () => new DOMRect(0, 0, 150, 100);
  run.getClientRects = () => [new DOMRect(0, 200, 100, 7.5), new DOMRect(0, 216, 100, 7.5)] as unknown as DOMRectList;
  expect(wordPrintGlyphBaseline(run)).toBeNull();
  expect(wordPrintGlyphBaseline(run, new DOMRect(0, 200, 5, 7.5))).toBeCloseTo(73.6);
  expect(wordPrintGlyphBaseline(run, new DOMRect(0, 216, 5, 7.5))).toBeCloseTo(105.6);
  expect(wordPrintGlyphBaseline(run, new DOMRect(0, 216, 5, 10))).toBeNull();
});

it('retains the leading line strut exactly once in each printed manual-break fragment', () => {
  const measured = Extension.create({
    name: 'testNativeMinimumLines',
    addProseMirrorPlugins: () => [
      new Plugin({
        props: {
          decorations(state) {
            return DecorationSet.create(
              state.doc,
              [0, 7].flatMap((from) => [
                Decoration.widget(
                  from + 1,
                  () => {
                    const strut = document.createElement('span');
                    strut.dataset.wordLineStrut = 'true';
                    strut.dataset.wordLineFrom = String(from);
                    strut.dataset.wordLineTo = String(from ? 12 : 7);
                    strut.style.cssText =
                      'display:inline-block;width:0;height:40px;vertical-align:top';
                    return strut;
                  },
                  { side: -1 },
                ),
                Decoration.inline(from + 1, from ? 13 : 8, {
                  class: 'word-page-fragment',
                  style: `--fragment-shift:${from ? 264 : 0}px`,
                }),
              ]),
            );
          },
        },
      }),
    ],
  });
  const editor = new Editor({
    extensions: [...wordExtensions(), measured],
    content:
      '<p style="font-size:10pt;font-family:Arial;line-height:30pt" data-word-line-rule="atLeast">before<span data-word-page-break="true"></span>after</p>',
  });
  try {
    const before = editor.getJSON();
    const editableHTML = editor.view.dom.innerHTML;
    expect(editor.view.dom.querySelector('.word-page-fragment')).not.toBeNull();
    const section = {
      id: 'word/document.xml#section:0',
      margins: { top: 600, right: 600, bottom: 600, left: 600, gutter: 0 },
      width: 6000,
      height: 4200,
    };
    const block = { from: 0, to: 14, section: 0, width: 320 };
    const input = { width: 400, blocks: [block], sections: [section] } as SurfaceInput;
    const plan = {
      width: 400,
      height: 584,
      blocks: [{ ...block, left: 40, top: 40 }],
      pages: [0, 304].map((top) => ({
        sectionId: section.id,
        left: 0,
        top,
        width: 400,
        height: 280,
      })),
      fragments: [
        {
          block: 0,
          from: 0,
          to: 7,
          page: 0,
          column: 0,
          shift: 0,
          shiftX: 0,
          left: 40,
          top: 40,
          height: 40,
        },
        {
          block: 0,
          from: 7,
          to: 12,
          page: 1,
          column: 0,
          shift: 264,
          shiftX: 0,
          left: 40,
          top: 344,
          height: 40,
        },
      ],
    } as WordFragmentPlan;
    const printed = wordFragmentPrint(editor.view, input, plan);
    const paragraphs = printed.querySelectorAll('p');
    expect(paragraphs).toHaveLength(2);
    for (const [index, p] of [...paragraphs].entries()) {
      const struts = p.querySelectorAll<HTMLElement>('[data-word-line-strut]');
      expect(struts).toHaveLength(1);
      expect(struts[0].dataset.wordLineFrom).toBe(String(index ? 7 : 0));
      expect(struts[0].style.height).toBe('40px');
      expect(Number(p.dataset.wordFragmentSourceTop)).toBe(index * 40);
    }
    expect(printed.textContent).toBe('beforeafter');
    expect(printed.querySelector('.word-page-fragment')).toBeNull();
    // A terminal caret can anchor the editable paragraph on a later page.
    // Moving that anchor must not move the text inside either print fragment.
    const anchored = wordFragmentPrint(editor.view, input, {
      ...plan,
      blocks: plan.blocks.map((b) => ({ ...b, top: b.top + 264 })),
      fragments: plan.fragments.map((f) => ({ ...f, shift: f.shift - 264 })),
    });
    expect([...anchored.querySelectorAll('p')].map((p) => p.dataset.wordFragmentSourceTop)).toEqual(
      ['0', '40'],
    );
    expect(anchored.textContent).toBe('beforeafter');
    // A continued paragraph retains its leading edge without restarting the
    // first-line indent on the next page. Print must not mutate the editor.
    const original = editor.view.nodeDOM(0) as HTMLElement;
    original.style.marginInlineStart = '12pt'; original.style.textIndent = '5pt';
    const indented = wordFragmentPrint(editor.view,
      { ...input, blocks: [{ ...block, indentStart: 16, width: 304 }] }, plan);
    const indentedParagraphs = [...indented.querySelectorAll('p')];
    expect(indentedParagraphs.map(p => p.style.left)).toEqual(['56px', '56px']);
    expect(indentedParagraphs.map(p => p.style.width)).toEqual(['304px', '304px']);
    expect(indentedParagraphs.map(p => parseFloat(p.style.textIndent))).toEqual([5, 0]);
    original.style.removeProperty('margin-inline-start'); original.style.removeProperty('text-indent');
    expect(editor.view.dom.innerHTML).toBe(editableHTML);
    expect(editor.getJSON()).toEqual(before);
    expect(editor.getHTML()).not.toContain('word-line-strut');
    expect(editor.getHTML()).not.toContain('word-fragment-source-top');
  } finally {
    editor.destroy();
  }
});
