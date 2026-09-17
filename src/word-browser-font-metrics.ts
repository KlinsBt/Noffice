/** Browser baseline measurements in the same CSS zoom and normal line box as
 * the editor. This is a disposable local probe, never document content. */
export function wordBrowserFontMetrics(
  style: CSSStyleDeclaration,
  size: number,
  natural: number,
  scale: number,
  cache: Map<string, { ascent: number; rangeAscent: number }>,
) {
  const id = `${[style.fontFamily, style.fontWeight, style.fontStyle, style.fontStretch].join('|')}|${size}|${scale}|${natural}`;
  let result = cache.get(id);
  if (!result) {
    const probe = document.createElement('span');
    probe.style.cssText =
      'all:initial;position:fixed;left:-100000px;top:0;visibility:hidden;white-space:pre;display:inline-block';
    probe.style.fontFamily = style.fontFamily;
    probe.style.fontWeight = style.fontWeight;
    probe.style.fontStyle = style.fontStyle;
    probe.style.fontStretch = style.fontStretch;
    probe.style.zoom = String(scale);
    probe.style.fontSize = '0';
    probe.style.lineHeight = '0';
    const inline = document.createElement('span');
    inline.style.fontSize = `${size}px`;
    inline.style.lineHeight = `${natural}px`;
    const text = document.createTextNode('Hg'),
      marker = document.createElement('span');
    marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline';
    inline.append(text);
    probe.append(inline, marker);
    document.body.append(probe);
    try {
      const range = document.createRange();
      range.selectNodeContents(text);
      const baseline = marker.getBoundingClientRect().top;
      result = {
        ascent: (baseline - probe.getBoundingClientRect().top) / scale,
        rangeAscent: (baseline - range.getBoundingClientRect().top) / scale,
      };
    } finally {
      probe.remove();
    }
    if (cache.size >= 256) cache.clear();
    cache.set(id, result);
  }
  return { ...result, natural };
}
