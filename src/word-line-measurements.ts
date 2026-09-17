import type { EditorView } from '@tiptap/pm/view';
import type { Node } from '@tiptap/pm/model';
import { wordHyphenText } from './word-hyphen';

export interface WordLine {
  /** UTF-16 offsets within the semantic paragraph; a hard break occupies one unit. */
  from: number;
  to: number;
  top: number;
  /** Advance to the next line, including trailing automatic leading. */
  height: number;
  /** Occupied height when this is the last line on a page or column. */
  fitHeight?: number;
  breakAfter?: 'page' | 'column';
}

/** Never choose a fragment boundary inside a grapheme, including across mark runs. */
export function paragraphGraphemes(node: Node): { text: string; boundaries: number[] } | null {
  let text = '',
    valid = true;
  node.forEach((child) => {
    if (child.isText) text += child.text;
    else if (child.type.name === 'wordTab') text += '\t';
    else if (child.type.name === 'wordHyphen' && wordHyphenText(child.attrs.kind) !== null)
      text += wordHyphenText(child.attrs.kind);
    else if (child.type.name === 'hardBreak') text += '\n';
    else if (child.type.name === 'wordPageBreak') text += '\f';
    else if (child.type.name === 'wordColumnBreak') text += '\u000e';
    else valid = false;
  });
  if (!valid || text.length > 20000 || typeof Intl.Segmenter !== 'function') return null;
  const boundaries = [
    ...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text),
  ].map((segment) => segment.index);
  boundaries.push(text.length);
  return { text, boundaries };
}

/** Measure existing DOM, after removing fragment offsets. No clone, text rewrite or
 * persistence change. Fixed-height line boxes are required for this first fragment path.
 * Unknown/nonuniform metrics return null instead of guessing a native line height.
 */
export function measureWordLines(view: EditorView, from: number): WordLine[] | null {
  const node = view.state.doc.nodeAt(from),
    dom = view.nodeDOM(from);
  if (!node || !(dom instanceof HTMLElement)) return null;
  const graphemes = paragraphGraphemes(node);
  if (!graphemes) return null;
  const withBreaks = (lines: WordLine[]) => {
    for (const line of lines) {
      const last = graphemes.text[line.to - 1];
      if (line.to > line.from && (last === '\f' || last === '\u000e'))
        line.breakAfter = last === '\f' ? 'page' : 'column';
    }
    return lines;
  };
  const style = getComputedStyle(dom),
    lineHeight = parseFloat(style.lineHeight),
    height = parseFloat(style.height),
    bounds = dom.getBoundingClientRect(),
    scale = bounds.width / parseFloat(style.width),
    count = Math.round(height / lineHeight);
  // The line-metric plugin already measures these semantic ranges and supplies
  // their native line advances. Consume that geometry instead of inferring a
  // uniform height from glyph rectangles (which vary with each run's font).
  const measured = [
    ...dom.querySelectorAll<HTMLElement>('[data-word-line-from][data-word-line-to]'),
  ];
  if (measured.length) {
    const boundaries = new Set(graphemes.boundaries);
    let top = 0;
    const lines = measured.map((strut) => {
      const line = {
        from: Number(strut.dataset.wordLineFrom),
        to: Number(strut.dataset.wordLineTo),
        top,
        height: parseFloat(strut.dataset.wordLineAdvance || strut.style.height),
        ...(strut.hasAttribute('data-word-line-fit-height')
          ? { fitHeight: Number(strut.dataset.wordLineFitHeight) }
          : {}),
      };
      top += line.height;
      return line;
    });
    return lines.length <= 10000 &&
      Math.abs(top - height) <= 0.2 &&
      lines.every(
        (line, index) =>
          Number.isSafeInteger(line.from) &&
          Number.isSafeInteger(line.to) &&
          Number.isFinite(line.height) &&
          line.height > 0 &&
          (line.fitHeight === undefined ||
            (Number.isFinite(line.fitHeight) &&
              line.fitHeight > 0 &&
              line.fitHeight <= line.height)) &&
          line.to >= line.from &&
          line.from === (index ? lines[index - 1].to : 0) &&
          boundaries.has(line.from) &&
          boundaries.has(line.to),
      ) &&
      lines.at(-1)!.to === node.content.size
      ? withBreaks(lines)
      : null;
  }
  if (dom.hasAttribute('data-word-mixed-leading') && Number.isFinite(height) && height > 0) {
    const after = parseFloat(style.getPropertyValue('--word-mixed-after'));
    return [
      {
        from: 0,
        to: node.content.size,
        top: 0,
        height,
        ...(after > 0 && after < height ? { fitHeight: height - after } : {}),
      },
    ];
  }
  if (
    !Number.isFinite(lineHeight) ||
    lineHeight <= 0 ||
    !Number.isFinite(scale) ||
    scale <= 0 ||
    count < 1 ||
    count > 10000 ||
    Math.abs(height - count * lineHeight) > 0.2
  )
    return null;
  const lines: WordLine[] = [];
  let previous = 0;
  const range = document.createRange();
  for (let i = 0; i < graphemes.boundaries.length - 1; i++) {
    const start = graphemes.boundaries[i],
      end = graphemes.boundaries[i + 1];
    const a = view.domAtPos(from + 1 + start, 1),
      b = view.domAtPos(from + 1 + end, -1);
    range.setStart(a.node, a.offset);
    range.setEnd(b.node, b.offset);
    const rects = [...range.getClientRects()];
    // A hard-break glyph is associated with the preceding line. Consecutive and
    // terminal hard breaks still produce empty lines in the model below.
    let index = ['\n', '\f', '\u000e'].includes(graphemes.text.slice(start, end))
      ? previous
      : Math.floor(
          ((rects[0]?.top ?? bounds.top) +
            (rects[0]?.height ?? lineHeight * scale) / 2 -
            bounds.top) /
            scale /
            lineHeight,
        );
    index = Math.max(previous, index);
    if (
      index >= count ||
      index < 0 ||
      (!rects.length && !['\n', '\f', '\u000e'].includes(graphemes.text.slice(start, end)))
    )
      return null;
    while (lines.length <= index)
      lines.push({ from: start, to: start, top: lines.length * lineHeight, height: lineHeight });
    lines[index].to = end;
    previous = ['\n', '\f', '\u000e'].includes(graphemes.text.slice(start, end))
      ? index + 1
      : index;
  }
  while (lines.length < count)
    lines.push({
      from: node.content.size,
      to: node.content.size,
      top: lines.length * lineHeight,
      height: lineHeight,
    });
  // Wrapping spaces are part of their original line's range; all semantic offsets
  // must remain covered once, in logical order.
  if (
    lines.at(-1)!.to !== node.content.size ||
    lines.some((l, i) => i > 0 && l.from !== lines[i - 1].to)
  )
    return null;
  return withBreaks(lines);
}
