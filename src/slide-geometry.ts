import type { SlideElement } from './model';

/** Resize from the lower-right handle, keeping the opposite corner fixed after rotation. */
export function resizeSlideElement(
  element: SlideElement,
  dx: number,
  dy: number,
  verticalScale: number,
) {
  const angle = ((element.rotation || 0) * Math.PI) / 180,
    cos = Math.cos(angle),
    sin = Math.sin(angle);
  const w = Math.max(30, Math.min(960, element.w + dx * cos + dy * sin));
  const h = Math.max(24, Math.min(540, element.h + (-dx * sin + dy * cos) / verticalScale));
  const dw = w - element.w,
    dh = (h - element.h) * verticalScale;
  return {
    w,
    h,
    x: element.x + (cos * dw - sin * dh - dw) / 2,
    y: element.y + (sin * dw + cos * dh - dh) / 2 / verticalScale,
  };
}
