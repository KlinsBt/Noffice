import { expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import native from '../tests/fixtures/native-powerpoint-arrange.json';
import { newSlide, textElement } from './model';
import { arrangeSlide, moveSlideSelection, type ArrangeCommand } from './slide-arrange';

const inputBytes = fs.readFileSync('tests/fixtures/arrange-oracle-cases.json');
const inputs = JSON.parse(inputBytes.toString());
it('binds every arrangement case to the exact native PowerPoint inputs', () => {
  expect(createHash('sha256').update(inputBytes).digest('hex')).toBe(native.inputSha256);
  expect(native.application).toBe('Microsoft PowerPoint');
  expect(native.cases.map((c) => c.name)).toEqual(inputs.map((c: { name: string }) => c.name));
});
it.each(native.cases)('matches native PowerPoint: $name', (c) => {
  const input = inputs.find((i: { name: string }) => i.name === c.name);
  const slide = {
    ...newSlide('blank'),
    elements: input.shapes.map(
      (s: { x: number; y: number; w: number; h: number; rotation: number }, i: number) =>
        textElement(`Object ${i}`, {
          ...s,
          x: (s.x * 960) / input.width,
          w: (s.w * 960) / input.width,
          y: (s.y * 540) / input.height,
          h: (s.h * 540) / input.height,
        }),
    ),
  };
  const arranged = arrangeSlide(
    slide,
    slide.elements.map((e: { id: string }) => e.id),
    input.command as ArrangeCommand,
    input.toSlide,
    input.width / input.height,
  );
  arranged.elements.forEach((el, i) => {
    expect((el.x * input.width) / 960).toBeCloseTo(c.positions[i].x, 2);
    expect((el.y * input.height) / 540).toBeCloseTo(c.positions[i].y, 2);
    expect({ ...el, x: 0, y: 0 }).toEqual({ ...slide.elements[i], x: 0, y: 0 });
  });
});
it('moves a selection without changing spacing, source identity, other objects or the original', () => {
  const first = textElement('First', { x: -5, y: 30, sourceShapeId: '2' }),
    second = textElement('Second', { x: 80, y: 150 }),
    untouched = textElement('Other');
  const slide = { ...newSlide('blank'), elements: [first, second, untouched] };
  const next = moveSlideSelection(slide, [first.id, second.id], -10, 25);
  expect(next.elements.map((e) => [e.x, e.y])).toEqual([
    [-15, 55],
    [70, 175],
    [untouched.x, untouched.y],
  ]);
  expect(next.elements[0].sourceShapeId).toBe('2');
  expect(next.elements[2]).toBe(untouched);
  expect(first.x).toBe(-5);
  expect(() => moveSlideSelection(slide, [first.id], NaN, 0)).toThrow();
  expect(() => arrangeSlide(slide, ['missing'], 'left', true)).toThrow();
  expect(arrangeSlide(slide, [first.id], 'left', false)).toBe(slide);
  expect(arrangeSlide(slide, [first.id, second.id], 'horizontal', false)).toBe(slide);
});
