import type { WordTabStop } from './word-tab-stops';

export interface WordTabLeaderFont {
  family: string;
  size: number;
  weight: string;
  style: string;
  stretch: string;
  normalRatio: number;
  letterSpacing: string;
  script?: 'superscript' | 'subscript';
}

/** Pinned Word/printer cell advances, independently checked against native
 * leader glyphs. These finite profiles are not arbitrary-text shaping or a
 * substitute for the remaining fonts, sizes, scripts and printer contracts. */
export function wordTabLeaderFont(leader: WordTabStop['leader'], font: WordTabLeaderFont) {
  const family = font.family.split(',')[0].replace(/["']/g, '').trim().toLowerCase();
  const regular = ['normal', '400'].includes(font.weight) && font.style === 'normal';
  if (leader === 'none' || !Number.isFinite(font.size) || !Number.isFinite(font.normalRatio)
    || !['normal', '100%'].includes(font.stretch)
    || (font.letterSpacing !== 'normal' && parseFloat(font.letterSpacing) !== 0)) return null;
  let pitches: readonly [number, number, number, number] | undefined;
  if (font.script !== undefined) {
    // Eight independently authored Arial10 superscript/subscript story cases.
    // The caller qualifies the nominal font and line context before deriving
    // the 6.48pt paint size. This does not qualify ordinary 6.48pt text.
    return ['superscript', 'subscript'].includes(font.script) && family === 'arial' && regular &&
      Math.abs(font.size - 6.48) < .0001 && Math.abs(font.normalRatio - 1.1499) <= .0001 && leader === 'dot'
      ? { glyph: '.', pitch: 1.8 } : null;
  }
  // Native mixed-size Grow/Shrink references independently measure these
  // dot advances. Other leader kinds at 9/11pt remain unmeasured.
  if (family === 'arial' && regular && leader === 'dot' && Math.abs(font.normalRatio - 1.1499) <= .0001) {
    if (Math.abs(font.size - 9) < .0001) return { glyph: '.', pitch: 2.52 };
    if (Math.abs(font.size - 11) < .0001) return { glyph: '.', pitch: 3.12 };
  }
  // The independently formatted Times20 story paragraph measures only dots.
  // Other leader kinds at this size remain unmeasured.
  if (family === 'times new roman' && regular && Math.abs(font.size - 20) < .0001
    && Math.abs(font.normalRatio - 1.1499) <= .0001 && leader === 'dot')
    return { glyph: '.', pitch: 5.04 };
  if (family === 'arial' && Math.abs(font.normalRatio - 1.1499) <= .0001) {
    const measuredStyle = regular || (['normal', '400'].includes(font.weight) && font.style === 'italic')
      || (['bold', '700'].includes(font.weight) && font.style === 'normal');
    if (Math.abs(font.size - 10) < .0001 && measuredStyle) pitches = [2.76, 3.36, 5.52, 3.36];
    else if (Math.abs(font.size - 20) < .0001 && regular) pitches = [5.52, 6.72, 11.16, 6.72];
  } else if (family === 'times new roman' && regular && Math.abs(font.size - 10) < .0001
    && Math.abs(font.normalRatio - 1.1499) <= .0001) pitches = [2.52, 3.24, 5.04, 3.36];
  else if (family === 'calibri' && regular && Math.abs(font.size - 11) < .0001
    && Math.abs(font.normalRatio - 1.2207) <= .0001) pitches = [2.76, 3.36, 5.52, 2.76];
  if (!pitches) return null;
  const index = { dot: 0, hyphen: 1, underscore: 2, heavy: 2, middleDot: 3 }[leader];
  return { glyph: ['.', '-', '_', '\u00b7'][index], pitch: pitches[index] };
}
