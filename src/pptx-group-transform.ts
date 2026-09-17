import type { SlideElement } from './model';

/** Affine transforms in source EMUs, composed from the outermost group inward. */
type Matrix = [number, number, number, number, number, number];
const identity = (): Matrix => [1, 0, 0, 1, 0, 0];
const child = (node: Element | undefined, name: string) =>
  node && Array.from(node.children).find((e) => e.localName === name);
const value = (node: Element | undefined, key: string, fallback = 0) => {
  const n = Number(node?.getAttribute(key) ?? fallback);
  if (!Number.isFinite(n)) throw Error('A group transform contains an invalid coordinate.');
  return n;
};
function multiply(a: Matrix, b: Matrix): Matrix {
  return [
    a[0] * b[0] + a[2] * b[1],
    a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3],
    a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4],
    a[1] * b[4] + a[3] * b[5] + a[5],
  ];
}
export function groupTransform(shape: Element) {
  const groups: Element[] = [];
  for (
    let parent = shape.parentElement;
    parent?.localName === 'grpSp';
    parent = parent.parentElement
  ) {
    groups.unshift(parent);
    if (groups.length > 100) throw Error('This presentation has excessively nested groups.');
  }
  if (!groups.length) return undefined;
  let matrix = identity();
  for (const group of groups) {
    const xfrm = child(child(group, 'grpSpPr'), 'xfrm');
    // The optional transform may be omitted when a container adds no transform.
    if (!xfrm) continue;
    const off = child(xfrm, 'off'),
      ext = child(xfrm, 'ext');
    const chOff = child(xfrm, 'chOff'),
      chExt = child(xfrm, 'chExt');
    const x = value(off, 'x'),
      y = value(off, 'y');
    const w = value(ext, 'cx'),
      h = value(ext, 'cy');
    const cw = value(chExt, 'cx'),
      ch = value(chExt, 'cy');
    if (Math.min(w, h, cw, ch) <= 0) throw Error('A group transform has zero or negative extents.');
    const angle = ((value(xfrm, 'rot') / 60000) * Math.PI) / 180;
    const fx = ['1', 'true'].includes(xfrm.getAttribute('flipH') || '') ? -1 : 1;
    const fy = ['1', 'true'].includes(xfrm.getAttribute('flipV') || '') ? -1 : 1;
    const a = (Math.cos(angle) * fx * w) / cw,
      b = (Math.sin(angle) * fx * w) / cw;
    const c = (-Math.sin(angle) * fy * h) / ch,
      d = (Math.cos(angle) * fy * h) / ch;
    const cx = value(chOff, 'x') + cw / 2,
      cy = value(chOff, 'y') + ch / 2;
    matrix = multiply(matrix, [
      a,
      b,
      c,
      d,
      x + w / 2 - a * cx - c * cy,
      y + h / 2 - b * cx - d * cy,
    ]);
  }
  if (
    matrix.some((n) => !Number.isFinite(n)) ||
    Math.abs(matrix[0] * matrix[3] - matrix[1] * matrix[2]) < 1e-12
  )
    throw Error('A group transform cannot be represented safely.');
  return matrix.map((n) => (n === 0 ? 0 : n)) as Matrix;
}

export function projectGroupElement(
  element: SlideElement,
  matrix: Matrix,
  width: number,
  height: number,
) {
  const local = {
    x: (element.x * width) / 960,
    y: (element.y * height) / 540,
    w: (element.w * width) / 960,
    h: (element.h * height) / 540,
    rotation: element.rotation || 0,
  };
  const r = (local.rotation * Math.PI) / 180;
  const axes = multiply(matrix, [Math.cos(r), Math.sin(r), -Math.sin(r), Math.cos(r), 0, 0]);
  const sx = Math.hypot(axes[0], axes[1]),
    sy = Math.hypot(axes[2], axes[3]);
  const projected = Math.abs(axes[0] * axes[2] + axes[1] * axes[3]) <= sx * sy * 1e-8;
  element.sourceGroupTransform = { matrix, local, projected, dimensions: [width, height] };
  // Non-orthogonal frames need a full affine scene renderer; retain source geometry explicitly.
  if (!projected) return;
  const cx = local.x + local.w / 2,
    cy = local.y + local.h / 2;
  const w = local.w * sx,
    h = local.h * sy;
  element.x = ((matrix[0] * cx + matrix[2] * cy + matrix[4] - w / 2) * 960) / width;
  element.y = ((matrix[1] * cx + matrix[3] * cy + matrix[5] - h / 2) * 540) / height;
  element.w = (w * 960) / width;
  element.h = (h * 540) / height;
  const reflected = axes[0] * axes[3] - axes[1] * axes[2] < 0;
  // Choose horizontal reflection so an unrotated flipped group keeps rotation=0.
  element.rotation =
    ((((Math.atan2(axes[1] * (reflected ? -1 : 1), axes[0] * (reflected ? -1 : 1)) * 180) /
      Math.PI) %
      360) +
      360) %
    360;
  if (element.type === 'image' && reflected) element.flipH = !element.flipH;
}

/** Move/resize orthogonal group frames while keeping child rotation/reflection unchanged. */
export function groupGeometryError(before: SlideElement, after: SlideElement) {
  if (!before.sourceGroupIds?.length) return undefined;
  const changed = (['x', 'y', 'w', 'h', 'rotation', 'flipH', 'flipV'] as const).some(
    (k) => before[k] !== after[k],
  );
  if (!changed) return undefined;
  const group = before.sourceGroupTransform;
  if (
    !group?.projected ||
    group.matrix.some((n) => !Number.isFinite(n)) ||
    Math.abs(group.matrix[0] * group.matrix[3] - group.matrix[1] * group.matrix[2]) < 1e-12 ||
    group.local.rotation !== 0 ||
    before.rotation !== after.rotation ||
    !!before.flipH !== !!after.flipH ||
    !!before.flipV !== !!after.flipV ||
    before.type === 'text'
  )
    return 'This edit requires unsupported grouped text, skew, rotation or reflection changes.';
  return undefined;
}
export function localGroupGeometry(
  before: SlideElement,
  after: SlideElement,
  width: number,
  height: number,
) {
  const error = groupGeometryError(before, after);
  if (error) throw Error(error);
  const matrix = before.sourceGroupTransform?.matrix || identity();
  const [a, b, c, d, tx, ty] = matrix;
  const sx = Math.hypot(a, b),
    sy = Math.hypot(c, d),
    determinant = a * d - b * c;
  const w = (after.w * width) / 960 / sx,
    h = (after.h * height) / 540 / sy;
  const cx = ((after.x + after.w / 2) * width) / 960 - tx;
  const cy = ((after.y + after.h / 2) * height) / 540 - ty;
  // The visible frame is centered on the transformed local center. Rotated or
  // reflected axes couple x/y with each other and with changes in width/height.
  const local = {
    x: (d * cx - c * cy) / determinant - w / 2,
    y: (a * cy - b * cx) / determinant - h / 2,
    w,
    h,
  };
  if (before.sourceGroupTransform) {
    const rounded = {
      x: Math.round(local.x),
      y: Math.round(local.y),
      w: Math.round(w),
      h: Math.round(h),
    };
    const centerX = rounded.x + rounded.w / 2,
      centerY = rounded.y + rounded.h / 2;
    const projected = {
      x: ((a * centerX + c * centerY + tx - (rounded.w * sx) / 2) * 960) / width,
      y: ((b * centerX + d * centerY + ty - (rounded.h * sy) / 2) * 540) / height,
      w: (rounded.w * sx * 960) / width,
      h: (rounded.h * sy * 540) / height,
    };
    for (const key of ['x', 'y', 'w', 'h'] as const) {
      if (!Number.isFinite(projected[key]) || Math.abs(projected[key] - after[key]) > 0.001)
        throw Error(
          'This group uses coarse coordinates. The edit cannot be exported accurately without rebuilding its coordinate system.',
        );
    }
  }
  return local;
}
