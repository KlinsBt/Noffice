import { outlineDashes, type SlideOutline } from './slide-outline';
export type DrawingPaint = { color: string; opacity: number };
export const drawingChild = (node: Element | null | undefined, name: string) =>
  node && Array.from(node.children).find((child) => child.localName === name);
const clamp = (n: number) => Math.max(0, Math.min(1, n));
const linear = (value: number) =>
  value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
const srgb = (value: number) =>
  value <= 0.0031308 ? value * 12.92 : 1.055 * value ** (1 / 2.4) - 0.055;
function hsl(rgb: number[]) {
  const max = Math.max(...rgb),
    min = Math.min(...rgb),
    delta = max - min,
    l = (max + min) / 2;
  const s = delta ? delta / (1 - Math.abs(2 * l - 1)) : 0;
  const h = !delta
    ? 0
    : max === rgb[0]
      ? ((rgb[1] - rgb[2]) / delta + 6) % 6
      : max === rgb[1]
        ? (rgb[2] - rgb[0]) / delta + 2
        : (rgb[0] - rgb[1]) / delta + 4;
  return [h / 6, s, l];
}
function rgb([h, s, l]: number[]) {
  const a = s * Math.min(l, 1 - l);
  return [0, 8, 4].map((n) => {
    const k = (n + h * 12) % 12;
    return l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1));
  });
}
/** Resolve a bounded subset of DrawingML solid colors, applying transforms in source order. */
export function drawingColor(
  node: Element | undefined | null,
  scheme: Record<string, string> = {},
  mapping: Record<string, string> = {},
  placeholder?: DrawingPaint,
): DrawingPaint | undefined {
  if (!node) return;
  let base: string | undefined;
  if (node.localName === 'srgbClr') base = node.getAttribute('val') || undefined;
  else if (node.localName === 'sysClr') base = node.getAttribute('lastClr') || undefined;
  else if (node.localName === 'schemeClr') {
    const key = node.getAttribute('val') || '';
    if (key === 'phClr') base = placeholder?.color.slice(1);
    else base = scheme[mapping[key] || key];
  }
  if (!base || !/^#?[a-f\d]{6}$/i.test(base)) return;
  const hex = base.replace('#', '');
  let channels = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255),
    opacity = node.getAttribute('val') === 'phClr' ? (placeholder?.opacity ?? 1) : 1;
  if (node.children.length > 64) return;
  for (const transform of Array.from(node.children)) {
    const amount = Number(transform.getAttribute('val')) / 100000;
    if (!Number.isFinite(amount)) return;
    if (transform.localName === 'tint')
      channels = channels.map((c) => srgb(clamp(linear(c) * amount + 1 - amount)));
    else if (transform.localName === 'shade')
      channels = channels.map((c) => srgb(clamp(linear(c) * amount)));
    else if (['lumMod', 'lumOff', 'satMod', 'satOff', 'lum', 'sat'].includes(transform.localName)) {
      const values = hsl(channels),
        axis = transform.localName.startsWith('lum') ? 2 : 1;
      values[axis] = clamp(
        transform.localName.endsWith('Mod')
          ? values[axis] * amount
          : transform.localName.endsWith('Off')
            ? values[axis] + amount
            : amount,
      );
      channels = rgb(values);
    } else if (transform.localName === 'alpha') opacity = clamp(amount);
    else if (transform.localName === 'alphaMod') opacity = clamp(opacity * amount);
    else if (transform.localName === 'alphaOff') opacity = clamp(opacity + amount);
    else return;
    channels = channels.map(clamp);
  }
  return {
    color:
      '#' +
      channels
        .map((c) =>
          Math.round(c * 255)
            .toString(16)
            .padStart(2, '0'),
        )
        .join(''),
    opacity,
  };
}

export function paintCSS(color: string, opacity = 1) {
  if (color === 'transparent' || opacity === 0) return 'transparent';
  return /^#[a-f\d]{6}$/i.test(color) && opacity < 1
    ? `rgba(${[1, 3, 5].map((i) => parseInt(color.slice(i, i + 2), 16)).join(', ')}, ${clamp(opacity)})`
    : color;
}

export function presentationPaints(
  theme: Document | null,
  master: Document | null,
  layout: Document | null,
  slide: Document,
) {
  const scheme: Record<string, string> = {};
  const colors = theme?.getElementsByTagNameNS('*', 'clrScheme')[0];
  for (const entry of Array.from(colors?.children || [])) {
    const color = drawingColor(entry.firstElementChild);
    if (color) scheme[entry.localName] = color.color.slice(1);
  }
  const mapping: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
  const sourceMap = master?.getElementsByTagNameNS('*', 'clrMap')[0];
  if (sourceMap)
    for (const attr of Array.from(sourceMap.attributes))
      if (!attr.name.startsWith('xmlns')) mapping[attr.localName] = attr.value;
  // A slide's masterClrMapping explicitly bypasses any layout override.
  const slideOverride = drawingChild(slide.documentElement, 'clrMapOvr');
  const override = slideOverride || (layout && drawingChild(layout.documentElement, 'clrMapOvr'));
  const custom = drawingChild(override, 'overrideClrMapping');
  if (custom)
    for (const attr of Array.from(custom.attributes))
      if (!attr.name.startsWith('xmlns')) mapping[attr.localName] = attr.value;
  const color = (node: Element | null | undefined, placeholder?: DrawingPaint) =>
    drawingColor(node, scheme, mapping, placeholder);
  function fill(
    node: Element | null | undefined,
    placeholder?: DrawingPaint,
  ): DrawingPaint | undefined {
    if (!node) return;
    if (node.localName === 'noFill') return { color: 'transparent', opacity: 0 };
    if (node.localName === 'solidFill') return color(node.firstElementChild, placeholder);
  }
  function reference(node: Element | null | undefined): DrawingPaint | undefined {
    if (!node) return;
    const index = Number(node.getAttribute('idx'));
    if (index === 0 || index === 1000) return { color: 'transparent', opacity: 0 };
    const list = theme?.getElementsByTagNameNS(
      '*',
      index >= 1001 ? 'bgFillStyleLst' : 'fillStyleLst',
    )[0];
    return fill(
      list?.children[index >= 1001 ? index - 1001 : index - 1],
      color(node.firstElementChild),
    );
  }
  function shape(shapes: (Element | null | undefined)[]): DrawingPaint | undefined {
    for (const item of shapes) {
      const properties = drawingChild(item, 'spPr');
      const choice =
        properties &&
        Array.from(properties.children).find((e) =>
          ['solidFill', 'noFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(
            e.localName,
          ),
        );
      if (choice) return fill(choice);
      const ref = drawingChild(drawingChild(item, 'style'), 'fillRef');
      if (ref) return reference(ref);
    }
  }
  function background(): DrawingPaint {
    for (const doc of [slide, layout, master]) {
      const bg = drawingChild(drawingChild(doc?.documentElement, 'cSld'), 'bg');
      if (bg) {
        const properties = drawingChild(bg, 'bgPr');
        const paint = properties
          ? fill(properties.firstElementChild)
          : reference(drawingChild(bg, 'bgRef'));
        return paint || { color: '#ffffff', opacity: 1 };
      }
    }
    return { color: '#ffffff', opacity: 1 };
  }
  function outline(
    shapes: (Element | null | undefined)[],
    slideWidth: number,
  ): SlideOutline | undefined {
    const sources: { node: Element; placeholder?: DrawingPaint }[] = [];
    for (const item of shapes) {
      const local = drawingChild(drawingChild(item, 'spPr'), 'ln');
      if (local) sources.push({ node: local });
      const ref = drawingChild(drawingChild(item, 'style'), 'lnRef');
      if (ref) {
        const idx = Number(ref.getAttribute('idx'));
        const styled = theme?.getElementsByTagNameNS('*', 'lnStyleLst')[0]?.children[idx - 1];
        if (styled) sources.push({ node: styled, placeholder: color(ref.firstElementChild) });
        break;
      }
    }
    let paint: DrawingPaint | undefined,
      width: number | undefined,
      dash: SlideOutline['dash'] | undefined;
    for (const source of sources) {
      const node = source.node;
      if (!paint) {
        const choice = Array.from(node.children).find((e) =>
          ['noFill', 'solidFill', 'gradFill', 'pattFill'].includes(e.localName),
        );
        if (choice) {
          paint = fill(choice, source.placeholder);
          if (!paint) return;
        }
      }
      if (width === undefined && node.hasAttribute('w'))
        width = (Number(node.getAttribute('w')) * 960) / slideWidth;
      if (!dash) {
        if (drawingChild(node, 'custDash')) return;
        const value = drawingChild(node, 'prstDash')?.getAttribute('val');
        if (value) {
          if (!(outlineDashes as readonly string[]).includes(value)) return;
          dash = value as SlideOutline['dash'];
        }
      }
    }
    if (!paint || !Number.isFinite(width ?? 0) || (width ?? 0) < 0 || (width ?? 0) > 1000) return;
    return { ...paint, width: width ?? 0, dash: dash ?? 'solid' };
  }
  return { color, shape, background, outline };
}
