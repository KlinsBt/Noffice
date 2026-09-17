import { wordPdfFontKey, type WordPdfSnapshot } from './word-pdf-model';

export interface WordPdfDecorationMetrics {
  unitsPerEm: number;
  underlinePosition: number;
  underlineThickness: number;
  strikePosition: number;
  strikeThickness: number;
}

/** Shared physical geometry for native-qualified screen and PDF decorations. */
export function wordDecorationGeometry(font: WordPdfDecorationMetrics, nominalSize: number, baseline: number, bit: number) {
  const step = .12, size = Math.round(nominalSize / step) * step;
  const position = bit === 1 ? font.underlinePosition : font.strikePosition;
  const thickness = bit === 1 ? font.underlineThickness : font.strikeThickness;
  return { y: baseline - Math.round(position * size / font.unitsPerEm / step) * step,
    height: Math.max(step, Math.round(thickness * size / font.unitsPerEm / step) * step) };
}

/** Read the exact embedded face, after the worker's font-directory validation.
 * These OpenType positions describe the top edge relative to the baseline. */
export function wordPdfDecorationMetrics(data: Uint8Array): WordPdfDecorationMetrics {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const table = (tag: number, minimum: number) => {
    if (data.length < 12) throw Error('Invalid PDF decoration font metrics.');
    const count = view.getUint16(4);
    if (count > 256 || 12 + count * 16 > data.length) throw Error('Invalid PDF decoration font metrics.');
    for (let i = 0; i < count; i++) {
      const entry = 12 + i * 16;
      if (view.getUint32(entry) !== tag) continue;
      const offset = view.getUint32(entry + 8), length = view.getUint32(entry + 12);
      if (length < minimum || offset + length > data.length) break;
      return offset;
    }
    throw Error('The PDF font lacks valid underline or strikethrough metrics.');
  };
  const head = table(0x68656164, 20), post = table(0x706f7374, 12), os2 = table(0x4f532f32, 30);
  const result = { unitsPerEm: view.getUint16(head + 18),
    underlinePosition: view.getInt16(post + 8), underlineThickness: view.getInt16(post + 10),
    strikePosition: view.getInt16(os2 + 28), strikeThickness: view.getInt16(os2 + 26) };
  if (result.unitsPerEm < 16 || result.unitsPerEm > 16384
    || result.underlineThickness <= 0 || result.strikeThickness <= 0)
    throw Error('The PDF font has invalid underline or strikethrough metrics.');
  return result;
}

/** Use Word's pinned 600dpi printer grid. The independent size/face matrix
 * validates vertical geometry; the shared layout supplies horizontal extents. */
export function wordPdfDecorationRules(
  page: WordPdfSnapshot['pages'][number], fonts: Map<string, WordPdfDecorationMetrics>,
) {
  const rules: { x: number; y: number; width: number; height: number; color: [number, number, number] }[] = [];
  for (const bit of [1, 2]) {
    let previous: { signature: string; rule: typeof rules[number] } | undefined;
    for (const glyph of page.decorations || []) {
      if (!(glyph.decoration & bit)) { previous = undefined; continue; }
      const key = wordPdfFontKey(glyph), font = fonts.get(key);
      if (!font) throw Error('The PDF decoration font metrics are unavailable.');
      const { y, height } = wordDecorationGeometry(font, glyph.size, glyph.y, bit);
      const signature = JSON.stringify([key, glyph.size, glyph.y, glyph.color, glyph.run]);
      const width = glyph.width!;
      // Printer-aligned character origins can differ from DOM range widths.
      // A semantic run's underline spans those advances (including kerning
      // gaps/overlaps); it must not become one disconnected bar per glyph.
      if (previous?.signature === signature && glyph.x >= previous.rule.x) {
        previous.rule.width = Math.max(previous.rule.width, glyph.x + width - previous.rule.x);
      } else if (width > 0) {
        const rule = { x: glyph.x, y, width, height, color: glyph.color };
        rules.push(rule); previous = { signature, rule };
      }
    }
  }
  return rules;
}
