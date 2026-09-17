import { expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import native from '../tests/fixtures/native-powerpoint-stack.json';
import { reorderStack, stackSlide, type StackCommand } from './slide-stack';
import { newSlide, textElement } from './model';
const bytes = fs.readFileSync('tests/fixtures/stack-oracle-cases.json');
const cases = JSON.parse(bytes.toString()) as {
  name: string;
  order: string[];
  selected: string[];
  command: StackCommand;
}[];
it('binds every layering case to the authored native inputs', () => {
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(native.inputSha256);
  expect(native.cases.map((c) => c.name)).toEqual(cases.map((c) => c.name));
});
it.each(native.cases)('matches native PowerPoint layering: $name', (c) => {
  const input = cases.find((row) => row.name === c.name)!;
  expect(reorderStack(input.order, input.selected, input.command)).toEqual(c.order);
});
it('keeps opaque source layers, source identities and relative selection order', () => {
  const a = textElement('A', { sourceShapeId: '2', sourceStackKey: 'source:2' });
  const b = textElement('B', { sourceShapeId: '4', sourceStackKey: 'source:4' });
  const slide = {
    ...newSlide('blank'),
    sourcePath: 'ppt/slides/slide1.xml',
    stackOrder: ['source:2', 'source:3', 'source:4'],
    elements: [a, b],
  };
  const next = stackSlide(slide, [a.id], 'forward');
  expect(next.stackOrder).toEqual(['source:3', 'source:2', 'source:4']);
  expect(next.elements).toEqual([a, b]);
  expect(slide.stackOrder).toEqual(['source:2', 'source:3', 'source:4']);
  expect(stackSlide(slide, [b.id], 'front')).toBe(slide);
  expect(() => stackSlide(slide, ['missing'], 'back')).toThrow('selection');
  expect(() => stackSlide({ ...slide, stackOrder: undefined }, [a.id], 'front')).toThrow('Reopen');
  expect(() =>
    stackSlide({ ...slide, elements: [{ ...a, sourceStackKey: 'source:9' }, b] }, [a.id], 'front'),
  ).toThrow('group');
});
