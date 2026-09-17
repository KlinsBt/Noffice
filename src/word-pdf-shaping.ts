import type { Font, TypeFeatures } from '@pdf-lib/fontkit';
import { wordPdfFontKey, type WordPdfSnapshot } from './word-pdf-model';

export function wordPdfFontFeatures(font: Font, value: number): TypeFeatures {
  // Fontkit adds default features to this object. Disable unrelated optional
  // substitutions explicitly and give each layout call its own copy.
  return { ...Object.fromEntries(font.availableFeatures.map((tag) => [tag, false])),
    kern: false, liga: !!(value & 1), clig: !!(value & 2), hlig: !!(value & 4),
    dlig: !!(value & 8), calt: !!(value & 16) };
}

/** Preserve measured cluster starts and semantic Unicode. A shaped glyph may
 * represent several characters; per-character PDF emission loses that glyph.
 * Context-dependent substitutions which cannot be encoded independently fail
 * explicitly until the PDF adapter can encode their complete shaped run. */
export function shapeWordPdfPage(page: WordPdfSnapshot['pages'][number], fonts: Map<string, Font>) {
  const output: typeof page.glyphs = [];
  for (let start = 0; start < page.glyphs.length;) {
    const first = page.glyphs[start];
    if (!first.features || first.run === undefined || first.text === '\u00a0' || first.literalHyphen) { output.push(first); start++; continue; }
    const key = wordPdfFontKey(first), font = fonts.get(key);
    if (!font) throw Error('The PDF shaping font is unavailable.');
    let end = start + 1;
    while (end < page.glyphs.length) {
      const next = page.glyphs[end];
      if (next.text === '\u00a0' || next.literalHyphen || next.run !== first.run || next.y !== first.y || next.size !== first.size
        || next.features !== first.features || wordPdfFontKey(next) !== key) break;
      end++;
    }
    const text = page.glyphs.slice(start, end).map((glyph) => glyph.text).join('');
    const options = wordPdfFontFeatures(font, first.features);
    const shaped = font.layout(text, { ...options });
    let offset = 0;
    for (const [index, glyph] of shaped.glyphs.entries()) {
      const cluster = String.fromCodePoint(...glyph.codePoints), position = shaped.positions[index];
      const encoded = font.layout(cluster, { ...options }).glyphs;
      if (!cluster || text.slice(offset, offset + cluster.length) !== cluster
        || !position || position.xOffset !== 0 || position.yOffset !== 0
        || encoded.length !== 1 || encoded[0].id !== glyph.id)
        throw Error('PDF export does not yet support this font shaping. The original DOCX is preserved.');
      output.push({ ...page.glyphs[start + offset], text: cluster });
      offset += cluster.length;
    }
    if (offset !== text.length) throw Error('The PDF font shaping lost document text.');
    start = end;
  }
  return { ...page, glyphs: output };
}
