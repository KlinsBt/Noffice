import {
  PDFArray,
  PDFDict,
  PDFName,
  PDFNumber,
  PDFOperator,
  PDFOperatorNames,
  PDFHexString,
  beginText,
  endText,
  setFontAndSize,
  setTextMatrix,
  pushGraphicsState,
  popGraphicsState,
  setFillingRgbColor,
  type PDFDocument,
  type PDFPage,
  type PDFFont,
} from 'pdf-lib';
import { wordPdfDrawFontKey, type WordPdfSnapshot } from './word-pdf-model';

export type WordPdfFullFontGlyphs = Map<string, Map<number, { text: string; width: number }>>;

/** Word paints NBSP through a separate ordinary-space text object. Retain the
 * original Unicode in the layout snapshot; canonicalize only PDF font mapping
 * so an aliased space glyph cannot corrupt other spaces' ToUnicode entries. */
export const wordPdfEncodedText = (text: string) => text === '\u00a0' ? ' ' : text;
// Fontkit hides U+00AD as default-ignorable when it first encounters that
// character in a subset. Select the measured minus glyph while the dedicated
// native WinAnsi dictionary retains codeAD and the original character map.
export const wordPdfPaintText = (text: string) => text === '\u00ad' ? '-' : wordPdfEncodedText(text);

/** Keep compatible text together, including authored trailing spaces. Native
 * paint changes start a new object; leading whitespace occupies its own object.
 * Every space is encoded, and TJ retains each measured origin after shaping. */
export function drawWordPdfText(
  document: PDFDocument,
  page: PDFPage,
  input: WordPdfSnapshot['pages'][number],
  fonts: Map<string, PDFFont>,
  fullFontGlyphs: WordPdfFullFontGlyphs,
) {
  const keys = new Map<string, ReturnType<PDFPage['node']['newFontDictionary']>>();
  for (let start = 0; start < input.glyphs.length;) {
    const first = input.glyphs[start];
    const fontKey = wordPdfDrawFontKey(first), font = fonts.get(fontKey)!;
    let key = keys.get(fontKey);
    if (!key) {
      key = page.node.newFontDictionary(font.name, font.ref);
      keys.set(fontKey, key);
    }
    const text = PDFArray.withContext(document.context);
    let end = start,
      previous = first;
    while (end < input.glyphs.length) {
      const glyph = input.glyphs[end];
      if (
        glyph.y !== first.y ||
        wordPdfDrawFontKey(glyph) !== fontKey ||
        glyph.size !== first.size ||
        (glyph.decoration ?? 0) !== (first.decoration ?? 0) ||
        (first.text === ' ' && glyph.text !== ' ') ||
        (glyph.text === '\u00a0') !== (first.text === '\u00a0') ||
        glyph.color.some((c, i) => c !== first.color[i])
      )
        break;
      if (end > start) {
        const expected =
          previous.x +
          (Math.round(font.widthOfTextAtSize(wordPdfPaintText(previous.text), 1000)) * first.size) / 1000;
        text.push(PDFNumber.of(((expected - glyph.x) * 1000) / first.size));
      }
      const encodedText = wordPdfEncodedText(glyph.text);
      const paintText = wordPdfPaintText(glyph.text);
      const encoded = font.encodeText(paintText), mapping = fullFontGlyphs.get(fontKey);
      if (mapping) {
        const bytes = encoded.asBytes(), width = Math.round(font.widthOfTextAtSize(paintText, 1000));
        if (bytes.length !== 2 || !Number.isFinite(width) || width < 0)
          throw Error('PDF export cannot encode this full-font glyph accurately. The original DOCX is preserved.');
        const cid = bytes[0] * 256 + bytes[1], previous = mapping.get(cid);
        if (previous && (previous.text !== encodedText || previous.width !== width))
          throw Error('PDF export cannot preserve this font\'s ambiguous text mapping. The original DOCX is preserved.');
        mapping.set(cid, { text: encodedText, width });
      }
      // Native Word uses WinAnsi code AD for a literal hyphen. Its dedicated
      // dictionary is finalized below over the intact TrueType font program.
      text.push(glyph.literalHyphen ? PDFHexString.of('AD') : encoded);
      previous = glyph;
      end++;
    }
    page.pushOperators(
      pushGraphicsState(),
      setFillingRgbColor(...first.color),
      beginText(),
      setFontAndSize(key, first.size),
      setTextMatrix(1, 0, 0, 1, first.x, input.height - first.y),
      PDFOperator.of(PDFOperatorNames.ShowTextAdjusted, [text]),
      endText(),
      popGraphicsState(),
    );
    start = end;
  }
}

/** Use integral CID widths and matching TJ adjustments. Some PDF readers
 * truncate fractional W entries; leaving those fractions in the font while
 * compensating from fontkit's full precision accumulates position drift. */
export async function finalizeWordPdfFonts(
  document: PDFDocument,
  fonts: Map<string, PDFFont>,
  fullFontGlyphs: WordPdfFullFontGlyphs,
) {
  for (const [key, font] of fonts) {
    await font.embed();
    const parent = document.context.lookup(font.ref, PDFDict);
    const descendants = parent.lookup(PDFName.of('DescendantFonts'), PDFArray);
    const cid = descendants.lookup(0, PDFDict);
    const used = fullFontGlyphs.get(key);
    if (key.endsWith('\0literal-hyphen')) {
      const glyphs = [...(used?.values() || [])];
      const descriptor = cid.lookup(PDFName.of('FontDescriptor'), PDFDict);
      if (glyphs.length !== 1 || glyphs[0].text !== '\u00ad' || !descriptor.has(PDFName.of('FontFile2')))
        throw Error('PDF export cannot preserve this font\'s literal hyphen. The original DOCX is preserved.');
      const flags = descriptor.lookup(PDFName.of('Flags'), PDFNumber).asNumber();
      descriptor.set(PDFName.of('Flags'), PDFNumber.of((flags & ~4) | 32));
      parent.set(PDFName.of('Subtype'), PDFName.of('TrueType'));
      parent.set(PDFName.of('Encoding'), PDFName.of('WinAnsiEncoding'));
      parent.set(PDFName.of('FontDescriptor'), cid.get(PDFName.of('FontDescriptor'))!);
      parent.set(PDFName.of('FirstChar'), PDFNumber.of(173));
      parent.set(PDFName.of('LastChar'), PDFNumber.of(173));
      parent.set(PDFName.of('Widths'), document.context.obj([glyphs[0].width]));
      parent.delete(PDFName.of('DescendantFonts')); parent.delete(PDFName.of('ToUnicode'));
      continue;
    }
    if (used?.size) {
      // Full embedding retains the font bytes. Its default pdf-lib character
      // cache omits substituted glyphs.
      // Describe actual emitted CIDs and their semantic Unicode/advance.
      const entries = [...used].sort(([a], [b]) => a - b);
      cid.set(PDFName.of('W'), document.context.obj(entries.flatMap(([code, glyph]) => [code, [glyph.width]])));
      const hex = (value: number) => value.toString(16).padStart(4, '0');
      const lines = [
        '/CIDInit /ProcSet findresource begin', '12 dict begin', 'begincmap',
        '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
        '/CMapName /Noffice-UTF16 def', '/CMapType 2 def',
        '1 begincodespacerange', '<0000> <ffff>', 'endcodespacerange',
      ];
      for (let start = 0; start < entries.length; start += 100) {
        const batch = entries.slice(start, start + 100);
        lines.push(`${batch.length} beginbfchar`);
        for (const [code, glyph] of batch)
          lines.push(`<${hex(code)}> <${glyph.text.split('').map(char => hex(char.charCodeAt(0))).join('')}>`);
        lines.push('endbfchar');
      }
      lines.push('endcmap', 'CMapName currentdict /CMap defineresource pop', 'end', 'end');
      parent.set(PDFName.of('ToUnicode'), document.context.register(document.context.flateStream(lines.join('\n'))));
      continue;
    }
    const widths = cid.lookup(PDFName.of('W'), PDFArray);
    for (let i = 1; i < widths.size(); i += 2) {
      const values = widths.lookup(i, PDFArray);
      for (let j = 0; j < values.size(); j++)
        values.set(j, PDFNumber.of(Math.round(values.lookup(j, PDFNumber).asNumber())));
    }
  }
}
