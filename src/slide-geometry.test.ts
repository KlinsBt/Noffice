import { expect, it } from 'vitest';
import { textElement } from './model';
import { resizeSlideElement } from './slide-geometry';
it.each([0, 30, 90, 270])(
  'keeps the opposite corner fixed while resizing at %s degrees on a 4:3 slide',
  (rotation) => {
    const el = textElement('Text', { x: 150, y: 160, w: 300, h: 120, rotation }),
      scale = 4 / 3;
    const angle = (rotation * Math.PI) / 180,
      c = Math.cos(angle),
      s = Math.sin(angle);
    const corner = (e: typeof el) => ({
      x: e.x + e.w / 2 - (c * e.w) / 2 + (s * e.h * scale) / 2,
      y: (e.y + e.h / 2) * scale - (s * e.w) / 2 - (c * e.h * scale) / 2,
    });
    const before = corner(el),
      changed = { ...el, ...resizeSlideElement(el, 40 * c - 20 * s, 40 * s + 20 * c, scale) },
      after = corner(changed);
    expect(changed.w).toBeCloseTo(340, 8);
    expect(changed.h).toBeCloseTo(135, 8);
    expect(after.x).toBeCloseTo(before.x, 8);
    expect(after.y).toBeCloseTo(before.y, 8);
  },
);
