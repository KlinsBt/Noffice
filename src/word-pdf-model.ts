/** A marked advance in semantic order. Tabs contribute paint without a glyph. */
export interface WordPdfDecorationSpan {
  x: number;
  y: number;
  size: number;
  family: string;
  face?: 'bold' | 'italic' | 'boldItalic';
  color: [number, number, number];
  run: number;
  /** Single underline (1), single strike (2), or both (3). */
  decoration: number;
  width: number;
}

/** Physical points, measured from the page's top-left. This is an ephemeral
 * rendering snapshot; it is never document content or an undo transaction. */
export interface WordPdfSnapshot {
  content: string;
  pages: {
    width: number;
    height: number;
    rules?: { x: number; y: number; width: number; height: number }[];
    decorations?: WordPdfDecorationSpan[];
    glyphs: {
      text: string;
      /** Qualified visible U+00AD from a literal Word text character. */
      literalHyphen?: true;
      /** Source-qualified list markers are paint, not document text. */
      listMarker?: true;
      x: number;
      y: number;
      size: number;
      family: string;
      face?: 'bold' | 'italic' | 'boldItalic';
      features?: number;
      /** Paint boundary; does not split a compatible shaping cluster. */
      decoration?: number;
      /** A compatible typography group within one paragraph, excluding tabs. */
      run?: number;
      color: [number, number, number];
    }[];
  }[];
}

export function wordPdfDrawFontKey(glyph: WordPdfSnapshot['pages'][number]['glyphs'][number]) {
  const font = wordPdfFontKey(glyph);
  const features = glyph.features ? `${font}\0features:${glyph.features}` : font;
  // Installed fonts alias literal U+00AD and U+002D to the same glyph ID.
  // Separate mappings preserve both characters without altering font bytes.
  return glyph.literalHyphen ? `${features}\0literal-hyphen` : features;
}

export function wordPdfFontKey(glyph: Pick<WordPdfSnapshot['pages'][number]['glyphs'][number], 'family' | 'face'>) {
  return glyph.face ? `${glyph.family}\0${glyph.face}` : glyph.family;
}

/** Printable Latin-1 has direct native glyph mappings. Discretionary soft
 * hyphens need a separate line-break contract; controls are never painted. */
export function wordPdfCharacterSupported(text: string): boolean {
  return /^[\x20-\x7e\u00a0-\u00ac\u00ae-\u00ff]$/.test(text);
}

export function validateWordPdfSnapshot(snapshot: WordPdfSnapshot) {
  if (
    !snapshot ||
    typeof snapshot.content !== 'string' ||
    snapshot.content.length > 10000000 ||
    !Array.isArray(snapshot.pages) ||
    !snapshot.pages.length ||
    snapshot.pages.length > 1000
  )
    throw Error('This document exceeds the supported PDF layout limits.');
  let count = 0,
    ruleCount = 0, decorationCount = 0;
  const fonts = new Set<string>();
  const variants = new Set<string>();
  for (const page of snapshot.pages) {
    if (
      ![page.width, page.height].every((n) => Number.isFinite(n) && n > 0 && n <= 1584) ||
      !Array.isArray(page.glyphs)
    )
      throw Error('Invalid PDF page geometry.');
    if (page.rules !== undefined && !Array.isArray(page.rules)) throw Error('Invalid PDF rules.');
    for (const rule of page.rules || []) {
      if (
        ++ruleCount > 5000 ||
        ![rule.x, rule.y, rule.width, rule.height].every(Number.isFinite) ||
        rule.width <= 0 ||
        rule.width > 2 ||
        rule.height <= 0 ||
        rule.height > page.height ||
        rule.x < 0 ||
        rule.y < 0 ||
        rule.x + rule.width > page.width ||
        rule.y + rule.height > page.height
      )
        throw Error('Invalid PDF rule geometry.');
    }
    for (const glyph of page.glyphs) {
      if (
        ++count > 50000 ||
        !(wordPdfCharacterSupported(glyph.text) || (glyph.text === '\u00ad' && glyph.literalHyphen === true) ||
          (glyph.text === '\u2022' && glyph.listMarker === true)) ||
        (glyph.listMarker !== undefined && (glyph.listMarker !== true || !/^[0-9.\u2022]$/.test(glyph.text))) ||
        (glyph.literalHyphen !== undefined && (glyph.literalHyphen !== true || glyph.text !== '\u00ad')) ||
        ![glyph.x, glyph.y, glyph.size].every(Number.isFinite) ||
        glyph.size <= 0 ||
        glyph.size > 400 ||
        Math.abs(glyph.x) > 3168 ||
        Math.abs(glyph.y) > 3168 ||
        typeof glyph.family !== 'string' ||
        !glyph.family ||
        glyph.family.length > 80 ||
        glyph.family.includes('\0') ||
        (glyph.face !== undefined && !['bold', 'italic', 'boldItalic'].includes(glyph.face)) ||
        (glyph.features !== undefined && (!Number.isInteger(glyph.features) || glyph.features < 0 || glyph.features > 31)) ||
        (glyph.decoration !== undefined && (!Number.isInteger(glyph.decoration) || glyph.decoration < 0 || glyph.decoration > 3)) ||
        (!!glyph.features && glyph.run === undefined) ||
        (glyph.run !== undefined && (!Number.isInteger(glyph.run) || glyph.run < 0 || glyph.run > 50000)) ||
        !Array.isArray(glyph.color) ||
        glyph.color.length !== 3 ||
        !glyph.color.every((n) => Number.isFinite(n) && n >= 0 && n <= 1)
      )
        throw Error('This text cannot yet be exported to PDF accurately.');
      fonts.add(wordPdfFontKey(glyph));
      variants.add(wordPdfDrawFontKey(glyph));
    }
    if (page.decorations !== undefined && !Array.isArray(page.decorations))
      throw Error('Invalid PDF decoration spans.');
    for (const span of page.decorations || []) {
      if (++decorationCount > 50000
        || ![span.x, span.y, span.size, span.width].every(Number.isFinite)
        || span.size <= 0 || span.size > 400 || span.width < 0 || span.width > 3168
        || Math.abs(span.x) > 3168 || Math.abs(span.y) > 3168
        || typeof span.family !== 'string' || !span.family || span.family.length > 80 || span.family.includes('\0')
        || (span.face !== undefined && !['bold', 'italic', 'boldItalic'].includes(span.face))
        || !Number.isInteger(span.run) || span.run < 0 || span.run > 50000
        || ![1, 2, 3].includes(span.decoration)
        || !Array.isArray(span.color) || span.color.length !== 3
        || !span.color.every(n => Number.isFinite(n) && n >= 0 && n <= 1))
        throw Error('Invalid PDF decoration span.');
      fonts.add(wordPdfFontKey(span));
    }
  }
  if (fonts.size > 8) throw Error('PDF export currently supports up to eight fonts per document.');
  if (variants.size > 32) throw Error('PDF export currently supports up to 32 font-feature or character-mapping variations per document.');
  return [...fonts];
}
