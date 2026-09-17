/** Neutral font-height signature shared with the finite native leader profiles.
 * These probes never enter an editor or its document/history. */
export function wordTabLeaderNormalRatio(style: CSSStyleDeclaration, cache: Map<string, number>) {
  const key = [style.fontFamily, style.fontWeight, style.fontStyle, style.fontStretch].join('|');
  const cached = cache.get(key);
  if (cached !== undefined) return cached;
  const probe = document.createElement('span');
  probe.style.cssText = 'all:initial;position:fixed;left:-100000px;top:0;visibility:hidden;pointer-events:none;white-space:pre;font-size:16384px;line-height:normal;display:inline-block';
  Object.assign(probe.style, { fontFamily: style.fontFamily, fontWeight: style.fontWeight,
    fontStyle: style.fontStyle, fontStretch: style.fontStretch });
  probe.textContent = 'Hg';document.body.append(probe);
  try {
    const ratio = probe.getBoundingClientRect().height / parseFloat(getComputedStyle(probe).fontSize);
    if (cache.size >= 256) cache.clear();
    if (Number.isFinite(ratio)) cache.set(key, ratio);
    return ratio;
  } finally { probe.remove(); }
}
