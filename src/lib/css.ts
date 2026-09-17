const unitless = new Set([
  'zoom',
  'fontWeight',
  'opacity',
  'zIndex',
  'flex',
  'flexGrow',
  'flexShrink',
  'order',
  'lineHeight',
]);
export function focusOnMount(node: HTMLElement) {
  queueMicrotask(() => node.focus());
}
export function css(styles: Record<string, unknown>): string {
  return Object.entries(styles)
    .filter(([, value]) => value !== undefined && value !== null)
    .map(
      ([key, value]) =>
        `${key.replace(/[A-Z]/g, (c) => '-' + c.toLowerCase())}:${typeof value === 'number' && !unitless.has(key) ? value + 'px' : value}`,
    )
    .join(';');
}
