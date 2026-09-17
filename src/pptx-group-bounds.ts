import { groupTransform } from './pptx-group-transform';

const direct = (node: Element, name: string) =>
  Array.from(node.children).find((child) => child.localName === name);
const number = (node: Element | undefined, name: string) =>
  node?.hasAttribute(name) ? Number(node.getAttribute(name)) : NaN;

/** Normalize affected group frames after all leaf edits, preserving their affine mapping.
 * Measured rectangle/ellipse children use native quadrant bounds; opaque rotated
 * frame types retain their containers. No descendants or payloads are rebuilt.
 */
export function normalizeGroupBounds(
  changedShapes: Element[],
  slideWidth: number,
  slideHeight: number,
) {
  const groups = new Set<Element>();
  for (const shape of changedShapes)
    for (let group = shape.parentElement; group?.localName === 'grpSp'; group = group.parentElement)
      groups.add(group);
  const mappings = new Map(
    [...groups].map((group) => {
      const child = Array.from(group.children).find((node) =>
        ['sp', 'pic', 'grpSp'].includes(node.localName),
      )!;
      return [group, { child, matrix: groupTransform(child)! }];
    }),
  );
  const checkpoints: { child: Element; matrix: number[]; points: number[][] }[] = [];
  const depth = (node: Element) => {
    let result = 0;
    for (let parent = node.parentElement; parent; parent = parent.parentElement) result++;
    return result;
  };
  for (const group of [...groups].sort((a, b) => depth(b) - depth(a))) {
    const properties = direct(group, 'grpSpPr');
    const transform = properties && direct(properties, 'xfrm');
    if (!transform) continue;
    const frames: { x: number; y: number; w: number; h: number }[] = [];
    let supported = true;
    for (const shape of Array.from(group.children)) {
      if (['nvGrpSpPr', 'grpSpPr', 'extLst'].includes(shape.localName)) continue;
      if (!['sp', 'pic', 'grpSp'].includes(shape.localName)) {
        supported = false;
        break;
      }
      const properties = direct(shape, shape.localName === 'grpSp' ? 'grpSpPr' : 'spPr');
      const frame = properties && direct(properties, 'xfrm');
      const rotation = Number(frame?.getAttribute('rot') || 0);
      if (
        !frame ||
        !Number.isSafeInteger(rotation) ||
        (rotation % 21600000 !== 0 &&
          (shape.localName !== 'sp' ||
            !['rect', 'ellipse'].includes(
              direct(properties!, 'prstGeom')?.getAttribute('prst') || '',
            ) ||
            Array.from(shape.getElementsByTagNameNS('*', 't')).some((t) => t.textContent)))
      ) {
        supported = false;
        break;
      }
      const off = direct(frame, 'off'),
        ext = direct(frame, 'ext');
      const rect = {
        x: number(off, 'x'),
        y: number(off, 'y'),
        w: number(ext, 'cx'),
        h: number(ext, 'cy'),
      };
      if (
        Object.values(rect).some((n) => !Number.isSafeInteger(n)) ||
        Math.min(rect.w, rect.h) <= 0
      ) {
        supported = false;
        break;
      }
      // Installed PowerPoint uses the unrotated rectangle, swapping dimensions
      // about its center for [45,135) modulo 180 degrees. This is deliberately
      // not a trigonometric union of rotated corners. Native boundary probes
      // cover both sides of 45/135/225/315 degrees across four group contexts.
      const angle = ((rotation % 10800000) + 10800000) % 10800000;
      if (angle >= 2700000 && angle < 8100000) {
        rect.x += (rect.w - rect.h) / 2;
        rect.y += (rect.h - rect.w) / 2;
        [rect.w, rect.h] = [rect.h, rect.w];
      }
      frames.push(rect);
    }
    if (!supported || !frames.length) continue;
    const off = direct(transform, 'off'),
      ext = direct(transform, 'ext');
    const chOff = direct(transform, 'chOff'),
      chExt = direct(transform, 'chExt');
    const x = number(off, 'x'),
      y = number(off, 'y'),
      w = number(ext, 'cx'),
      h = number(ext, 'cy');
    const cx = number(chOff, 'x'),
      cy = number(chOff, 'y');
    const cw = number(chExt, 'cx'),
      ch = number(chExt, 'cy');
    if (![x, y, w, h, cx, cy, cw, ch].every(Number.isSafeInteger) || Math.min(w, h, cw, ch) <= 0)
      continue;
    const bounds = frames.reduce(
      (b, f) => [
        Math.min(b[0], f.x),
        Math.min(b[1], f.y),
        Math.max(b[2], f.x + f.w),
        Math.max(b[3], f.y + f.h),
      ],
      [Infinity, Infinity, -Infinity, -Infinity],
    );
    const [left, top, right, bottom] = bounds;
    const width = right - left,
      height = bottom - top;
    checkpoints.push({
      ...mappings.get(group)!,
      points: [
        [left, top],
        [right, top],
        [left, bottom],
        [right, bottom],
      ],
    });
    const angle = (Number(transform.getAttribute('rot') || 0) * Math.PI) / 10800000;
    const fx = ['1', 'true'].includes(transform.getAttribute('flipH') || '') ? -1 : 1;
    const fy = ['1', 'true'].includes(transform.getAttribute('flipV') || '') ? -1 : 1;
    const dx = ((left + width / 2 - cx - cw / 2) * w * fx) / cw;
    const dy = ((top + height / 2 - cy - ch / 2) * h * fy) / ch;
    const nextW = (width * w) / cw,
      nextH = (height * h) / ch;
    const nextX = x + w / 2 + Math.cos(angle) * dx - Math.sin(angle) * dy - nextW / 2;
    const nextY = y + h / 2 + Math.sin(angle) * dx + Math.cos(angle) * dy - nextH / 2;
    const values = [nextX, nextY, nextW, nextH, left, top, width, height].map(Math.round);
    if (!values.every(Number.isSafeInteger) || [2, 3, 6, 7].some((i) => values[i] <= 0))
      throw Error('The edited group bounds cannot be represented safely.');
    for (const [index, node] of [off!, ext!, chOff!, chExt!].entries()) {
      const keys = index % 2 ? ['cx', 'cy'] : ['x', 'y'];
      keys.forEach((key, axis) => node.setAttribute(key, String(values[index * 2 + axis])));
    }
  }
  // Integer group coordinates can magnify rounding through nested scales. Reject
  // an inaccurate output before returning a download; the retained source is intact.
  for (const { child, matrix: before, points } of checkpoints) {
    const after = groupTransform(child)!;
    for (const [x, y] of points)
      for (const axis of [0, 1]) {
        const difference =
          (after[axis] - before[axis]) * x +
          (after[axis + 2] - before[axis + 2]) * y +
          after[axis + 4] -
          before[axis + 4];
        const units = axis ? slideHeight / 540 : slideWidth / 960;
        if (!Number.isFinite(difference) || Math.abs(difference) / units > 0.001)
          throw Error(
            'This group uses coarse coordinates. Its bounds cannot be normalized accurately without rebuilding its coordinate system.',
          );
      }
  }
}
