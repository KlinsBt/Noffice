import { PDFDocument, rgb } from 'pdf-lib';
import fontkit, { type Font } from '@pdf-lib/fontkit';
import decompress from 'woff2-encoder/decompress';
import { validateWordPdfSnapshot, wordPdfFontKey, wordPdfDrawFontKey, type WordPdfSnapshot } from './word-pdf-model';
import { drawWordPdfText, finalizeWordPdfFonts, wordPdfEncodedText, type WordPdfFullFontGlyphs } from './word-pdf-text';
import { shapeWordPdfPage, wordPdfFontFeatures } from './word-pdf-shaping';
import { wordPdfDecorationMetrics, wordPdfDecorationRules, type WordPdfDecorationMetrics } from './word-pdf-decorations';

type Input = { snapshot: WordPdfSnapshot; fonts: { family: string; data: ArrayBuffer }[] };

/** Validate table extents and embedding permissions before passing font data
 * to a parser. The worker deadline also bounds malformed-font computation. */
function validateFont(data: Uint8Array) {
  if (data.byteLength < 12 || data.byteLength > 16 * 1024 * 1024)
    throw Error('The PDF font is invalid or too large.');
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (![0x00010000, 0x4f54544f].includes(view.getUint32(0)))
    throw Error('PDF export needs an individual TrueType or OpenType font.');
  const count = view.getUint16(4);
  let subset = true;
  if (!count || count > 256 || 12 + count * 16 > data.length)
    throw Error('Invalid font directory.');
  for (let i = 0; i < count; i++) {
    const entry = 12 + i * 16,
      offset = view.getUint32(entry + 8),
      length = view.getUint32(entry + 12);
    if (offset + length > data.length) throw Error('Invalid font table.');
    if (view.getUint32(entry) === 0x4f532f32) {
      if (length < 10) throw Error('Invalid font embedding permissions.');
      const permissions = view.getUint16(offset + 8);
      if (permissions & 0x0202)
        throw Error('This font does not permit embedding its outlines in PDF.');
      if (permissions & 0x0100) subset = false;
    }
  }
  return subset;
}

self.onmessage = async (event: MessageEvent<Input>) => {
  try {
    const { snapshot, fonts } = event.data;
    const families = validateWordPdfSnapshot(snapshot);
    if (
      !Array.isArray(fonts) ||
      fonts.length !== families.length ||
      fonts.length > 8 ||
      fonts.reduce((n, f) => n + f.data.byteLength, 0) > 32 * 1024 * 1024
    )
      throw Error('The PDF font input exceeds its limits.');
    const document = await PDFDocument.create();
    document.registerFontkit(fontkit);
    document.setProducer('Noffice');
    document.setCreator('Noffice');
    const embedded = new Map<string, Awaited<ReturnType<typeof document.embedFont>>>();
    const shapingFonts = new Map<string, Font>();
    const decorationFonts = new Map<string, WordPdfDecorationMetrics>();
    const fullFontGlyphs: WordPdfFullFontGlyphs = new Map();
    for (const family of families) {
      const input = fonts.find((f) => f.family === family)?.data;
      if (
        !(input instanceof ArrayBuffer) ||
        input.byteLength < 12 ||
        input.byteLength > 10 * 1024 * 1024
      )
        throw Error(`The font ${family} is unavailable for PDF export.`);
      let data = new Uint8Array(input);
      if (new DataView(input).getUint32(0) === 0x774f4632) {
        // Reject declared expansion before invoking the WOFF2 decoder.
        if (input.byteLength < 48 || new DataView(input).getUint32(16) > 16 * 1024 * 1024)
          throw Error('The compressed font exceeds its supported size.');
        data = new Uint8Array(await decompress(data));
      }
      const subset = validateFont(data);
      if (snapshot.pages.some((page) => page.decorations?.some((span) =>
        wordPdfFontKey(span) === family)))
        decorationFonts.set(family, wordPdfDecorationMetrics(data));
      const shaping = fontkit.create(data);
      shapingFonts.set(family, shaping);
      const literalHyphens = snapshot.pages.some(page => page.glyphs.some(glyph =>
        glyph.literalHyphen && wordPdfFontKey(glyph) === family));
      if (literalHyphens && (new DataView(data.buffer, data.byteOffset, data.byteLength).getUint32(0) !== 0x00010000 ||
        shaping.glyphForCodePoint(0xad).id !== shaping.glyphForCodePoint(0x2d).id))
        throw Error('PDF export cannot yet preserve this font\'s literal hyphen. The original DOCX is preserved.');
      const supported = new Set(shaping.characterSet);
      if (
        snapshot.pages.some((page) =>
          page.glyphs.some(
            (glyph) => wordPdfFontKey(glyph) === family && !supported.has(wordPdfEncodedText(glyph.text).codePointAt(0)!),
          ),
        )
      )
        throw Error(`The font ${family} does not contain all text needed for PDF export.`);
      const variants = new Map(snapshot.pages.flatMap((page) => page.glyphs)
        .filter((glyph) => wordPdfFontKey(glyph) === family)
        .map((glyph) => [wordPdfDrawFontKey(glyph), glyph.features || 0]));
      for (const [key, features] of variants) {
        embedded.set(key, await document.embedFont(data, {
          // Native WinAnsi AD needs the font's original character map. A
          // minus-only subset lacks that map and paints a missing-glyph box.
          subset: subset && !key.endsWith('\0literal-hyphen'), features: wordPdfFontFeatures(shaping, features),
        }));
        if (!subset || key.endsWith('\0literal-hyphen')) fullFontGlyphs.set(key, new Map());
      }
    }
    for (const input of snapshot.pages) {
      const page = document.addPage([input.width, input.height]);
      for (const rule of input.rules || [])
        page.drawRectangle({
          x: rule.x,
          y: input.height - rule.y - rule.height,
          width: rule.width,
          height: rule.height,
          color: rgb(0, 0, 0),
        });
      const shaped = shapeWordPdfPage(input, shapingFonts);
      drawWordPdfText(document, page, shaped, embedded, fullFontGlyphs);
      for (const rule of wordPdfDecorationRules(input, decorationFonts))
        page.drawRectangle({ x: rule.x, y: input.height - rule.y - rule.height,
          width: rule.width, height: rule.height, color: rgb(...rule.color) });
    }
    await finalizeWordPdfFonts(document, embedded, fullFontGlyphs);
    const bytes = await document.save();
    if (bytes.byteLength > 32 * 1024 * 1024)
      throw Error('The generated PDF exceeds its supported size.');
    self.postMessage({ bytes }, { transfer: [bytes.buffer] });
  } catch (error) {
    self.postMessage({ error: error instanceof Error ? error.message : 'PDF export failed.' });
  }
};
