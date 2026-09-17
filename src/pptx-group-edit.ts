import type { SlideElement } from './model';
import { groupGeometryError } from './pptx-group-transform';

/** Native Size-pane edits retain the rotated upper-left corner of eligible
 * grouped children. Explicit model frames and pointer gestures stay separate.
 */
export function groupSizeField(
  before: SlideElement,
  key: 'w' | 'h',
  value: number,
): Partial<SlideElement> {
  const source = before.sourceGroupTransform;
  if (!source?.projected || source.local.rotation !== 0 || !before.rotation)
    return { [key]: value };
  const angle = (before.rotation * Math.PI) / 180;
  const dw = key === 'w' ? value - before.w : 0;
  const dh = key === 'h' ? value - before.h : 0;
  return {
    [key]: value,
    x: before.x + ((Math.cos(angle) - 1) * dw - Math.sin(angle) * dh) / 2,
    y: before.y + (Math.sin(angle) * dw + (Math.cos(angle) - 1) * dh) / 2,
  };
}

/** Visible size of one retained child-coordinate unit; independent of subsequent edits. */
export function coarseGroupGrid(element: SlideElement) {
  const source = element.sourceGroupTransform;
  if (!source?.dimensions) return undefined;
  const [width, height] = source.dimensions,
    [a, b, c, d] = source.matrix;
  const x = (Math.hypot(a, b) * 960) / width,
    y = (Math.hypot(c, d) * 540) / height;
  return Math.max(x, y) > 0.001 ? { x, y } : undefined;
}

/** Native UI commits eligible coarse children to their retained integer coordinate grid.
 * Keep the source mapping immutable, so history/reload never change the grid's origin.
 */
export function quantizeGroupEdit(before: SlideElement, after: SlideElement): SlideElement {
  const grid = coarseGroupGrid(before);
  if (!grid || !(['x', 'y', 'w', 'h'] as const).some((k) => before[k] !== after[k])) return after;
  const error = groupGeometryError(before, after);
  if (error) throw Error(error);
  const source = before.sourceGroupTransform!;
  const [a, b, c, d, tx, ty] = source.matrix,
    [width, height] = source.dimensions!;
  if (!source.projected || a <= 0 || d <= 0 || Math.abs(b) > 1e-8 || Math.abs(c) > 1e-8)
    throw Error(
      'Coarse coordinates in rotated or reflected groups cannot be edited accurately yet.',
    );
  const originX = (tx * 960) / width,
    originY = (ty * 540) / height;
  const values = {
    x: originX + Math.round((after.x - originX) / grid.x) * grid.x,
    y: originY + Math.round((after.y - originY) / grid.y) * grid.y,
    w: Math.round(after.w / grid.x) * grid.x,
    h: Math.round(after.h / grid.y) * grid.y,
  };
  if (
    Object.values(values).some((v) => !Number.isFinite(v) || v < 0 || v > 2000) ||
    values.w < 10 ||
    values.h < 10
  )
    throw Error('This size or position cannot be represented on the group coordinate grid.');
  // Avoid a storage revision/history event for a native no-op or floating-point noise.
  for (const key of ['x', 'y', 'w', 'h'] as const)
    if (Math.abs(values[key] - before[key]) < 1e-9) values[key] = before[key];
  return { ...after, ...values };
}
