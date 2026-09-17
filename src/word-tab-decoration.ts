import { wordDecorationGeometry } from './word-pdf-decorations';
import type { WordTabLeaderFont } from './word-tab-leader-font';

/** Numeric metrics of the pinned Arial/Calibri faces, independently checked
 * by the 256 native vertical controls. No font outlines are distributed here.
 * Unknown descriptors retain their ordinary editor fallback. */
export function wordTabDecoration(font: WordTabLeaderFont, marks: number) {
  const family = font.family.split(',')[0].trim().replace(/^["']|["']$/g, '').toLowerCase();
  if (!['arial', 'calibri'].includes(family) || ![1, 2, 3].includes(marks)
    || !Number.isFinite(font.size) || font.size < 1 || font.size > 400 || font.script
    || !['normal', '400', 'bold', '700'].includes(font.weight)
    || !['normal', 'italic'].includes(font.style) || !['normal', '100%'].includes(font.stretch)
    || (font.letterSpacing !== 'normal' && parseFloat(font.letterSpacing) !== 0)
    || !Number.isFinite(font.normalRatio)
    || Math.abs(font.normalRatio - (family === 'arial' ? 1.1499 : 1.2207)) > .0001) return null;
  const bold = ['bold', '700'].includes(font.weight);
  const metrics = family === 'arial'
    ? { unitsPerEm: 2048, underlinePosition: -217, underlineThickness: bold ? 215 : 150, strikePosition: 530, strikeThickness: 102 }
    : { unitsPerEm: 2048, underlinePosition: -232, underlineThickness: 134, strikePosition: 512, strikeThickness: bold ? 186 : 134 };
  const rules = [1, 2].filter(bit => marks & bit)
    .map(bit => wordDecorationGeometry(metrics, font.size, 0, bit)).sort((a, b) => a.y - b.y);
  const top = rules[0].y / .75, height = (rules.at(-1)!.y + rules.at(-1)!.height) / .75 - top;
  // Rasterize whole printer cells before scaling their coordinate system.
  // A subpixel-height CSS background can otherwise clip the last ink band.
  const paintHeight = Math.round(height / .16);
  const stops: string[] = [];
  let previous = 0;
  for (const rule of rules) {
    const start = Math.round((rule.y / .75 - top) / .16), end = start + Math.round(rule.height / .12);
    if (start > previous) stops.push(`transparent ${previous}px`, `transparent ${start}px`);
    stops.push(`currentColor ${start}px`, `currentColor ${end}px`); previous = end;
  }
  return { top, height, paintHeight, background: `linear-gradient(to bottom, ${stops.join(', ')})`, rules };
}
