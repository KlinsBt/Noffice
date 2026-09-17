import type { Slide, SlideElement } from './model';

export type StackCommand = 'front' | 'back' | 'forward' | 'backward';
export const stackKey = (element: SlideElement) =>
  element.sourceStackKey ||
  (element.sourceShapeId ? `source:${element.sourceShapeId}` : `new:${element.id}`);

/** Back-to-front order. Selection order must not change the selected shapes' relative order. */
export function reorderStack(
  order: readonly string[],
  selected: readonly string[],
  command: StackCommand,
): string[] {
  const chosen = new Set(selected);
  if (new Set(order).size !== order.length || selected.some((id) => !order.includes(id)))
    throw Error('The object stacking selection is no longer valid.');
  if (command === 'front')
    return [...order.filter((id) => !chosen.has(id)), ...order.filter((id) => chosen.has(id))];
  if (command === 'back')
    return [...order.filter((id) => chosen.has(id)), ...order.filter((id) => !chosen.has(id))];
  if (!['forward', 'backward'].includes(command)) throw Error('Unknown stacking command.');
  const next = [...order];
  if (command === 'forward') {
    for (let i = next.length - 2; i >= 0; i--)
      if (chosen.has(next[i]) && !chosen.has(next[i + 1]))
        [next[i], next[i + 1]] = [next[i + 1], next[i]];
  } else {
    for (let i = 1; i < next.length; i++)
      if (chosen.has(next[i]) && !chosen.has(next[i - 1]))
        [next[i], next[i - 1]] = [next[i - 1], next[i]];
  }
  return next;
}

export function stackSlide(slide: Slide, ids: readonly string[], command: StackCommand): Slide {
  const selection = new Set(ids);
  const chosen = slide.elements.filter((el) => selection.has(el.id));
  if (chosen.length !== selection.size) throw Error('The selection is no longer on this slide.');
  if (!chosen.length) return slide;
  if (slide.sourcePath && !slide.stackOrder)
    throw Error('Reopen this presentation to load its original object stacking order.');
  if (chosen.some((el) => el.sourceStackKey && el.sourceStackKey !== `source:${el.sourceShapeId}`))
    throw Error('Reordering individual grouped objects requires group-aware editing.');
  const keys = new Set(slide.elements.map(stackKey));
  // Unknown imported objects remain real stacking layers; deleted new objects do not.
  const order = (slide.stackOrder || []).filter((key) => !key.startsWith('new:') || keys.has(key));
  for (const el of slide.elements) if (!order.includes(stackKey(el))) order.push(stackKey(el));
  const next = reorderStack(order, chosen.map(stackKey), command);
  if (next.every((key, i) => key === order[i])) return slide;
  const positions = new Map(next.map((key, i) => [key, i]));
  return {
    ...slide,
    stackOrder: next,
    elements: [...slide.elements].sort(
      (a, b) => positions.get(stackKey(a))! - positions.get(stackKey(b))!,
    ),
  };
}
