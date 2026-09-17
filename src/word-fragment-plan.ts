import type { WordLine } from './word-line-measurements';
import type { SurfaceInput, BlockMeasurement, WordSurfacePlan } from './word-section-surfaces';
import { continuesSectionPage } from './word-section-surfaces';
import { wordStoryPageBounds } from './word-story-layout';

const fitHeight = (line: WordLine) => line.fitHeight ?? line.height;

export interface ParagraphFlow extends BlockMeasurement {
  lines: WordLine[];
  keepLines: boolean;
  widowControl: boolean;
}
export interface WordFragment {
  block: number;
  from: number;
  to: number;
  page: number;
  shift: number;
  column: number;
  shiftX: number;
  left: number;
  top: number;
  height: number;
}
export interface WordFragmentPlan extends WordSurfacePlan {
  fragments: WordFragment[];
}

/** Short sections already have complete block placements. Represent each whole
 * paragraph as a fragment so PDF export uses the same physical page geometry. */
export function fragmentWordSurfaces(
  plan: WordSurfacePlan,
  measurements: BlockMeasurement[],
): WordFragmentPlan {
  if (plan.blocks.length !== measurements.length) throw Error('Incomplete page measurements.');
  return {
    ...plan,
    fragments: plan.blocks.map((b, block) => ({
      block,
      from: 0,
      to: b.to - b.from - 2,
      page: b.section,
      column: 0,
      shift: 0,
      shiftX: 0,
      left: b.left,
      top: b.top + measurements[block].before,
      height: measurements[block].height,
    })),
  };
}

/** Page breaking over measured lines. Paragraphs/marks/source IDs remain semantic;
 * derived fragments contain offsets only and are never exported as new paragraphs.
 */
export function planWordFragments(
  input: SurfaceInput,
  measurements: ParagraphFlow[],
  gap = 24,
): WordFragmentPlan | null {
  const limits = new Map<number, { page: number; height: number }>();
  let plan = allocateWordFragments(input, measurements, gap, limits);
  if (!plan) return null;
  // Word balances the preceding section's final columns at a compatible
  // continuous transition. The document's final section fills sequentially.
  // Replay the same keep/widow allocator while finding the smallest usable
  // height, so balancing cannot choose a different set of break constraints.
  for (let section = 0; section + 1 < input.sections.length; section++) {
    if (!input.sections[section].columns || !continuesSectionPage(input, section + 1)) continue;
    const fragments = plan.fragments.filter((f) => input.blocks[f.block].section === section);
    const page = fragments.at(-1)!.page;
    const s = input.sections[section];
    let low = 0,
      high = (s.height! - s.margins.top! - s.margins.bottom!) / 15;
    for (let pass = 0; pass < 18 && high - low > 1 / 128; pass++) {
      const height = (low + high) / 2;
      limits.set(section, { page, height });
      const trial = allocateWordFragments(input, measurements, gap, limits);
      if (trial) high = height;
      else low = height;
    }
    limits.set(section, { page, height: Math.ceil(high * 64) / 64 });
    plan = allocateWordFragments(input, measurements, gap, limits);
    if (!plan) return null;
  }
  return plan;
}

function allocateWordFragments(
  input: SurfaceInput,
  measurements: ParagraphFlow[],
  gap: number,
  limits: Map<number, { page: number; height: number }>,
): WordFragmentPlan | null {
  if (measurements.length !== input.blocks.length || !Number.isFinite(gap) || gap < 0) return null;
  const result: WordFragmentPlan = {
    width: input.width,
    height: 0,
    pages: [],
    blocks: [],
    fragments: [],
    separators: [],
  };
  let block = 0;
  // An even-start first section begins at logical page 2 without a leading
  // physical blank. Subsequent parity breaks and story slots use that number.
  const firstPageNumber = input.sections[0]?.start === 'evenPage' ? 2 : 1;
  let continuation: { y: number; after: number; top: number } | undefined;
  for (let section = 0; section < input.sections.length; section++) {
    const s = input.sections[section],
      width = s.width! / 15,
      height = s.height! / 15,
      left = (input.width - width) / 2,
      columns = s.columns?.widths.length ?? 1;
    let bodyTop = s.margins.top! / 15,
      bodyBottom = height - s.margins.bottom! / 15,
      capacity = bodyBottom - bodyTop;
    const columnLeft = (index: number) =>
      s.columns
        ? s.columns.widths.slice(0, index).reduce((n, w, i) => n + w + s.columns!.spaces[i], 0) / 15
        : 0;
    let y = bodyTop,
      previousAfter = 0,
      pageTop = result.height,
      column = 0,
      columnTop = bodyTop,
      maximumY = bodyTop;
    let sectionPage: number | undefined;
    const sealColumns = () => {
      if (sectionPage === undefined || !s.columns?.separator || column === 0) return;
      for (let boundary = 0; boundary < column; boundary++)
        result.separators!.push({
          page: sectionPage,
          left:
            left +
            s.margins.left! / 15 +
            columnLeft(boundary) +
            (s.columns.widths[boundary] + s.columns.spaces[boundary] / 2) / 15,
          top: pageTop + columnTop,
          height: maximumY - columnTop,
        });
    };
    const limit = limits.get(section);
    const bottom = () =>
      limit?.page === result.pages.length - 1
        ? Math.min(bodyBottom, columnTop + limit.height)
        : bodyBottom;
    const page = (blank = false) => {
      if (result.pages.length >= 500 || (limit && result.pages.length > limit.page)) return false;
      sealColumns();
      pageTop = result.height;
      bodyTop = s.margins.top! / 15;
      bodyBottom = height - s.margins.bottom! / 15;
      capacity = bodyBottom - bodyTop;
      const bounds =
        input.stories && !blank
          ? wordStoryPageBounds(
              s,
              input.stories,
              result.pages.length + firstPageNumber,
              sectionPage === undefined,
            )
          : undefined;
      if (bounds === null) return false;
      if (bounds) {
        bodyTop = bounds.top;
        bodyBottom = bounds.bottom;
        capacity = bodyBottom - bodyTop;
      }
      result.pages.push({
        sectionId: s.id,
        left,
        top: pageTop,
        width,
        height,
        bodyTop,
        bodyBottom,
        ...(bounds ? { stories: bounds.stories } : {}),
        ...(blank ? { blank: true } : {}),
      });
      // Parity padding is outside the incoming section's first body page.
      sectionPage = blank ? undefined : result.pages.length - 1;
      result.height += height + gap;
      y = bodyTop;
      columnTop = bodyTop;
      maximumY = bodyTop;
      column = 0;
      previousAfter = 0;
      return true;
    };
    if (continuation && continuesSectionPage(input, section)) {
      y = continuation.y;
      previousAfter = continuation.after;
      pageTop = continuation.top;
      columnTop = y;
      maximumY = y;
      sectionPage = result.pages.length - 1;
      const currentPage = result.pages[sectionPage];
      if (currentPage.bodyTop !== undefined && currentPage.bodyBottom !== undefined) {
        bodyTop = currentPage.bodyTop;
        bodyBottom = currentPage.bodyBottom;
        capacity = bodyBottom - bodyTop;
      }
    } else {
      if (
        section > 0 &&
        ['oddPage', 'evenPage'].includes(s.start) &&
        (result.pages.length + firstPageNumber) % 2 !== (s.start === 'oddPage' ? 1 : 0)
      ) {
        // Native parity blanks use the incoming section's paper dimensions.
        // They have no body fragment or semantic paragraph to save or edit.
        if (!page(true)) return null;
      }
      if (!page()) return null;
    }
    const advance = () => {
      maximumY = Math.max(maximumY, y);
      if (column + 1 >= columns) return page();
      column++;
      y = columnTop;
      previousAfter = 0;
      return true;
    };
    while (block < input.blocks.length && input.blocks[block].section === section) {
      const b = input.blocks[block],
        m = measurements[block],
        lines = m.lines;
      if (
        !lines.length ||
        ![m.height, m.before, m.after].every(Number.isFinite) ||
        m.height <= 0 ||
        m.before < 0 ||
        m.after < 0 ||
        Math.abs(m.height - lines.at(-1)!.top - lines.at(-1)!.height) > 0.2 ||
        lines.some(
          (l, i) =>
            ![l.from, l.to, l.top, l.height].every(Number.isFinite) ||
            l.height <= 0 ||
            l.height > capacity ||
            !Number.isFinite(fitHeight(l)) ||
            fitHeight(l) <= 0 ||
            fitHeight(l) > l.height ||
            l.to < l.from ||
            (i === 0
              ? l.from !== 0 || l.top !== 0
              : l.from !== lines[i - 1].to ||
                Math.abs(l.top - lines[i - 1].top - lines[i - 1].height) > 0.2),
        ) ||
        lines.at(-1)!.to !== b.to - b.from - 2
      )
        return null;
      // Word paints a trailing empty section paragraph on the preceding line
      // at a compatible continuous transition. It contributes no extra line
      // or paragraph spacing. A section consisting only of that paragraph
      // still occupies its normal line. Keep the semantic paragraph/caret and
      // its measured fragment so typing can immediately restore ordinary flow.
      if (
        b.to - b.from === 2 &&
        block > 0 &&
        input.blocks[block - 1].section === section &&
        input.blocks[block + 1]?.section === section + 1 &&
        !s.columns &&
        !b.pageBreakBefore &&
        !b.keepNext &&
        !input.blocks[block - 1].flowBreaks &&
        continuesSectionPage(input, section + 1)
      ) {
        const prior = result.fragments.at(-1)!;
        const priorLines = measurements[block - 1].lines;
        const first = priorLines.find((line) => line.from === prior.from)!;
        const top = prior.top + priorLines.at(-1)!.top - first.top;
        result.blocks.push({ ...b, left: prior.left, top: top - m.before });
        result.fragments.push({
          block,
          from: 0,
          to: 0,
          page: prior.page,
          column: prior.column,
          left: prior.left,
          top,
          height: m.height,
          shift: 0,
          shiftX: 0,
        });
        block++;
        continue;
      }
      if (b.pageBreakBefore && (y > bodyTop || column > 0) && !page()) return null;
      y += Math.max(previousAfter, m.before);
      // Manual breaks divide the paragraph's keep/widow constraints into
      // portions without splitting its semantic model or source identity.
      const portions: { from: number; to: number; height: number }[] = [];
      let portionFrom = 0;
      for (let i = 0; i < lines.length; i++)
        if (lines[i].breakAfter || i === lines.length - 1) {
          portions.push({
            from: portionFrom,
            to: i + 1,
            height: lines[i].top + fitHeight(lines[i]) - lines[portionFrom].top,
          });
          portionFrom = i + 1;
        }
      let start = 0,
        portionIndex = 0,
        origin: number | undefined,
        originX = 0;
      while (start < lines.length) {
        const portion = portions[portionIndex];
        const oversizedKeep = m.keepLines && portion.height > capacity;
        if (start === portion.from) {
          if (oversizedKeep && (y > bodyTop || column > 0) && !page()) return null;
          const whole = m.keepLines || (m.widowControl && portion.to - portion.from <= 3);
          while (whole && portion.height <= capacity && y + portion.height > bottom() + 0.01)
            if (!advance()) return null;
        }
        let end = start,
          used = 0;
        // Extra automatic leading advances the next line, but can extend past
        // a page/column's body when the final line's natural box still fits.
        while (end < portion.to && y + used + fitHeight(lines[end]) <= bottom() + 0.01) {
          used += lines[end].height;
          end++;
        }
        if (end === lines.length && b.keepNext) {
          // Keep-with-next binds the last line to the required first lines of
          // the next paragraph, rather than necessarily moving both paragraphs.
          // Include a full keep-together paragraph and chained one-line headings.
          let required = 0,
            trailing = 0,
            after = m.after;
          for (let next = block + 1; next < input.blocks.length; next++) {
            const nextBlock = input.blocks[next],
              nextFlow = measurements[next];
            if (nextBlock.section !== section || nextBlock.pageBreakBefore) break;
            const forcedEnd = nextFlow.lines.findIndex((line) => line.breakAfter);
            const first = forcedEnd < 0 ? nextFlow.lines.length : forcedEnd + 1;
            const firstHeight =
              nextFlow.lines[first - 1].top + fitHeight(nextFlow.lines[first - 1]);
            const take =
              nextFlow.keepLines && firstHeight <= capacity
                ? first
                : Math.min(nextFlow.widowControl ? 2 : 1, first);
            required +=
              Math.max(after, nextFlow.before) +
              nextFlow.lines.slice(0, take).reduce((sum, line) => sum + line.height, 0);
            const last = nextFlow.lines[take - 1];
            trailing = last.height - fitHeight(last);
            if (
              required - trailing > capacity ||
              take < nextFlow.lines.length ||
              !nextBlock.keepNext
            )
              break;
            after = nextFlow.after;
          }
          required -= trailing;
          if (
            required > 0 &&
            y + used + required > bottom() + 0.01 &&
            lines[end - 1].height + required <= capacity + 0.01
          ) {
            end--;
            used -= lines[end].height;
          }
        }
        if (end < portion.to && m.widowControl) {
          if (end - start < 2 && y > columnTop) end = start;
          else if (portion.to - end === 1 && end - start > 2) {
            end--;
            used -= lines[end].height;
          } else if (portion.to - end === 1 && end - start === 2 && y > columnTop) end = start;
        }
        if (end === start) {
          if (!advance()) return null;
          continue;
        }
        if (origin === undefined) {
          origin = pageTop + y;
          originX = columnLeft(column);
          result.blocks.push({
            ...b,
            left: left + s.margins.left! / 15 + originX,
            top: origin - m.before,
          });
        }
        result.fragments.push({
          block,
          from: lines[start].from,
          to: lines[end - 1].to,
          page: result.pages.length - 1,
          shift: pageTop + y - origin - lines[start].top,
          column,
          shiftX: columnLeft(column) - originX,
          left: left + s.margins.left! / 15 + columnLeft(column),
          top: pageTop + y,
          height: used,
        });
        y += used;
        maximumY = Math.max(maximumY, y);
        start = end;
        if (start === portion.to) portionIndex++;
        // Word's oversized kept paragraph continues on physical pages rather
        // than splitting across the section's other columns.
        if (
          start < lines.length &&
          !((oversizedKeep && start < portion.to) || lines[end - 1].breakAfter === 'page'
            ? page()
            : advance())
        )
          return null;
      }
      if (b.pageBreakParagraph && lines.at(-1)?.breakAfter === 'page' && !page()) return null;
      previousAfter = m.after;
      block++;
    }
    sealColumns();
    continuation = { y: maximumY, after: previousAfter, top: pageTop };
  }
  result.height -= gap;
  // Anchor the actual paragraph to its empty terminal line. Its trailing BR and
  // native caret then need no synthetic text/widget; earlier inline fragments
  // are positioned relative to that anchor, including on preceding pages.
  for (const terminal of result.fragments) {
    if (terminal.from !== terminal.to || (terminal.shift === 0 && terminal.shiftX === 0)) continue;
    const shift = terminal.shift;
    const shiftX = terminal.shiftX;
    result.blocks[terminal.block].top += shift;
    result.blocks[terminal.block].left += shiftX;
    for (const fragment of result.fragments)
      if (fragment.block === terminal.block) {
        fragment.shift -= shift;
        fragment.shiftX -= shiftX;
      }
  }
  return result;
}
