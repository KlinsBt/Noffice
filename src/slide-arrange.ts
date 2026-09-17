import type { Slide, SlideElement } from './model';

export type ArrangeCommand =
  'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom' | 'horizontal' | 'vertical';

function bounds(element: SlideElement, verticalScale: number, distribute: boolean) {
  // Native PowerPoint distribution uses quadrant-snapped extents; alignment uses rotated extents.
  const rotation = distribute
    ? Math.round((element.rotation || 0) / 90) * 90
    : element.rotation || 0;
  const radians = (rotation * Math.PI) / 180;
  const width =
    Math.abs(element.w * Math.cos(radians)) +
    Math.abs(element.h * verticalScale * Math.sin(radians));
  const height =
    (Math.abs(element.w * Math.sin(radians)) +
      Math.abs(element.h * verticalScale * Math.cos(radians))) /
    verticalScale;
  return {
    x: element.x + (element.w - width) / 2,
    y: element.y + (element.h - height) / 2,
    w: width,
    h: height,
  };
}

/** Retain source identities, object order and non-geometric properties. */
export function arrangeSlide(
  slide: Slide,
  ids: readonly string[],
  command: ArrangeCommand,
  toSlide: boolean,
  aspectRatio = 16 / 9,
): Slide {
  if (!Number.isFinite(aspectRatio) || aspectRatio <= 0) throw Error('Invalid slide proportions.');
  const selection = new Set(ids);
  const chosen = slide.elements.filter((el) => selection.has(el.id));
  if (chosen.length !== selection.size) throw Error('The selection is no longer on this slide.');
  const distribute = command === 'horizontal' || command === 'vertical';
  if (chosen.length < (toSlide ? 1 : distribute ? 3 : 2)) return slide;
  const horizontal = ['left', 'center', 'right', 'horizontal'].includes(command);
  const axis = horizontal ? 'x' : 'y',
    size = horizontal ? 'w' : 'h';
  const verticalScale = 960 / aspectRatio / 540;
  const entries = chosen.map((el) => ({ el, box: bounds(el, verticalScale, distribute) }));
  const positions = new Map<string, number>();
  if (distribute) {
    entries.sort((a, b) => a.box[axis] + a.box[size] / 2 - b.box[axis] - b.box[size] / 2);
    const start = toSlide ? 0 : entries[0].box[axis];
    const last = entries.at(-1)!;
    const end = toSlide ? (horizontal ? 960 : 540) : last.box[axis] + last.box[size];
    const gap =
      (end - start - entries.reduce((sum, entry) => sum + entry.box[size], 0)) /
      (toSlide ? entries.length + 1 : entries.length - 1);
    let cursor = start + (toSlide ? gap : 0);
    for (const { el, box } of entries) {
      positions.set(el.id, cursor + el[axis] - box[axis]);
      cursor += box[size] + gap;
    }
  } else {
    const start = toSlide ? 0 : Math.min(...entries.map(({ box }) => box[axis]));
    const end = toSlide
      ? horizontal
        ? 960
        : 540
      : Math.max(...entries.map(({ box }) => box[axis] + box[size]));
    const factor =
      command === 'center' || command === 'middle'
        ? 0.5
        : command === 'right' || command === 'bottom'
          ? 1
          : 0;
    for (const { el, box } of entries)
      positions.set(el.id, start + (end - start - box[size]) * factor + el[axis] - box[axis]);
  }
  return {
    ...slide,
    elements: slide.elements.map((el) =>
      positions.has(el.id) ? { ...el, [axis]: positions.get(el.id)! } : el,
    ),
  };
}

export function moveSlideSelection(
  slide: Slide,
  ids: readonly string[],
  dx: number,
  dy: number,
): Slide {
  if (!Number.isFinite(dx) || !Number.isFinite(dy)) throw Error('Invalid object movement.');
  const selection = new Set(ids);
  return {
    ...slide,
    elements: slide.elements.map((el) =>
      selection.has(el.id) ? { ...el, x: el.x + dx, y: el.y + dy } : el,
    ),
  };
}
