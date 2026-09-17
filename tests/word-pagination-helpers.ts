import type { Locator, Page } from '@playwright/test';

export async function waitWordLayout(editor: Locator) {
  await editor.evaluate(async (root) => {
    await document.fonts.ready;
    let previous = '',
      stable = 0;
    for (let frame = 0; frame < 30 && stable < 3; frame++) {
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
      const signature = root.innerHTML;
      stable = signature === previous ? stable + 1 : 0;
      previous = signature;
    }
    if (stable < 3) throw Error('Word page layout did not settle before pointer input.');
  });
}

export async function clickWordText(page: Page, editor: Locator, word: string) {
  await waitWordLayout(editor);
  const point = await editor.evaluate((root, word) => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const text = walker.currentNode as Text,
        index = text.data.indexOf(word);
      if (index < 0 || text.parentElement?.closest('[contenteditable=false]')) continue;
      text.parentElement!.scrollIntoView({ block: 'center' });
      const range = document.createRange();
      range.setStart(text, index);
      range.setEnd(text, index + 1);
      const rect = range.getBoundingClientRect();
      return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
    }
    throw Error(`Missing word ${word}`);
  }, word);
  await page.mouse.click(point.x, point.y);
}

export const wordLineOrigins = async (editor: Locator) =>
  editor.locator('p').evaluateAll((paragraphs) =>
    paragraphs.map((p) => {
      const output: { page: number; x: number; y: number }[] = [];
      let start = true;
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
      let node: Node | null;
      while ((node = walker.nextNode())) {
        const el = node instanceof Element ? node : node.parentElement;
        if (node instanceof Element && node.classList.contains('word-reflow-line')) start = true;
        if (
          node instanceof Element &&
          node.matches('[data-word-page-break], [data-word-column-break]')
        ) {
          start = true;
          continue;
        }
        if (el?.closest('[contenteditable=false]')) continue;
        if (node instanceof HTMLBRElement) {
          start = true;
          continue;
        }
        if (node.nodeType !== Node.TEXT_NODE || !node.textContent || !start) continue;
        const range = document.createRange();
        range.setStart(node, 0);
        range.setEnd(node, 1);
        const rect = range.getBoundingClientRect();
        const pages = [...document.querySelectorAll<HTMLElement>('.section-page')];
        const index = pages.findIndex((page) => {
          const b = page.getBoundingClientRect();
          return rect.top + rect.height / 2 >= b.top && rect.top + rect.height / 2 < b.bottom;
        });
        if (index < 0) throw Error('A printed line lies outside the page surfaces.');
        const page = pages[index],
          bounds = page.getBoundingClientRect();
        const scale = bounds.width / parseFloat(getComputedStyle(page).width);
        output.push({
          page: index + 1,
          x: ((rect.left - bounds.left) * 0.75) / scale,
          y: ((rect.top - bounds.top) * 0.75) / scale,
        });
        start = false;
      }
      return output;
    }),
  );

export const wordPlacements = async (editor: Locator) =>
  editor.locator('p').evaluateAll((paragraphs) =>
    paragraphs.map((p) => {
      const characters: { text: string; page: number }[] = [];
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
      let n: Node | null;
      while ((n = walker.nextNode())) {
        const el = n instanceof Element ? n : n.parentElement;
        const flow =
          n instanceof Element && n.matches('[data-word-page-break], [data-word-column-break]');
        if (!flow && el?.closest('[contenteditable=false]')) continue;
        const text = flow
          ? el!.hasAttribute('data-word-page-break')
            ? '\f'
            : '\u000e'
          : n.nodeType === Node.TEXT_NODE
            ? n.textContent!
            : el?.tagName === 'BR' && !el.classList.contains('ProseMirror-trailingBreak')
              ? '\n'
              : '';
        for (let i = 0; i < text.length; i++) {
          const range = document.createRange();
          if (n.nodeType === Node.TEXT_NODE) {
            range.setStart(n, i);
            range.setEnd(n, i + 1);
          } else range.selectNode(n);
          const r = range.getBoundingClientRect(),
            y = r.top + r.height / 2;
          // Short imported documents retain their single source paper instead
          // of adding overflow fragments. Both are physical page surfaces.
          const fragments = [...document.querySelectorAll('.section-page')];
          const papers = fragments.length ? fragments : [...document.querySelectorAll('.paper')];
          const rectangles = papers.map((e) => e.getBoundingClientRect());
          const page = rectangles.findIndex((rect) => y >= rect.top && y < rect.bottom) + 1;
          characters.push({ text: text[i], page });
        }
      }
      return characters;
    }),
  );
