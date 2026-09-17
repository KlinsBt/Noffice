/** Stored paragraph spacing is independent of the font's physical line box. */
export type WordParagraphSpace = { unit: 'points' | 'lines'; value: number } | { unit: 'auto' };
export type WordSpaceSide = 'before' | 'after';

export function wordParagraphSpace(value: unknown): WordParagraphSpace | null {
  if (typeof value === 'string') {
    if (value.length > 100) return null;
    try { value = JSON.parse(value); } catch { return null; }
  }
  if (!value || typeof value !== 'object') return null;
  const space = value as Record<string, unknown>;
  if (space.unit === 'auto') return { unit: 'auto' };
  if ((space.unit !== 'points' && space.unit !== 'lines') ||
      !Number.isInteger(space.value) || Number(space.value) < 0 ||
      Number(space.value) > (space.unit === 'points' ? 31680 : 32767)) return null;
  return { unit: space.unit, value: Number(space.value) };
}

/** Word's COM inputs are Single precision; persisted units are integral. */
export function wordParagraphSpaceInput(unit: WordParagraphSpace['unit'], input: unknown): WordParagraphSpace | null {
  if (unit === 'auto') return { unit };
  if ((typeof input !== 'number' && typeof input !== 'string') || String(input).trim() === '') return null;
  const value = Number(input);
  if (!Number.isFinite(value) || value < 0) return null;
  return wordParagraphSpace({ unit, value: Math.round(Math.fround(value) * (unit === 'points' ? 20 : 100)) });
}

export function wordParagraphSpaceTwips(space: WordParagraphSpace): number {
  // The pinned native controls render Auto as 14pt for the measured ordinary
  // paragraphs, including the legacy HTML-auto-spacing compatibility flag.
  // Contextual suppression is a view operation, never a change to this value.
  if (space.unit === 'auto') return 280;
  return space.unit === 'points' ? space.value : Math.min(31680, Math.trunc(space.value * 12 / 5));
}

export function wordParagraphPointSpace(css: unknown): WordParagraphSpace | null {
  if (css == null || css === '') return { unit: 'points', value: 0 };
  if (typeof css !== 'string' || !/^\d+(?:\.\d+)?(?:pt|px)$/.test(css)) return null;
  return wordParagraphSpace({ unit: 'points', value: Math.round(parseFloat(css) * (css.endsWith('pt') ? 20 : 15)) });
}

export function wordParagraphSpaceFromAttrs(attrs: Record<string, unknown>, side: WordSpaceSide) {
  const suffix = side === 'before' ? 'Before' : 'After';
  return wordParagraphSpace(attrs['paragraphSpace' + suffix]) ?? wordParagraphPointSpace(attrs['space' + suffix]);
}

export function wordParagraphSpaceAttributes(side: WordSpaceSide, space: WordParagraphSpace) {
  const suffix = side === 'before' ? 'Before' : 'After';
  return {
    ['space' + suffix]: `${wordParagraphSpaceTwips(space) / 20}pt`,
    ['paragraphSpace' + suffix]: space.unit === 'points' ? null : space,
  };
}

/** Explicit mode switches must cancel inherited modes as well as direct ones. */
export function wordParagraphSpaceXml(side: WordSpaceSide, space: WordParagraphSpace) {
  return {
    [side]: String(space.unit === 'auto' ? 0 : wordParagraphSpaceTwips(space)),
    [side + 'Lines']: String(space.unit === 'lines' ? space.value : 0),
    [side + 'Autospacing']: space.unit === 'auto' ? '1' : '0',
  };
}
