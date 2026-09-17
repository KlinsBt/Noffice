import { expect, it } from 'vitest';
import { PDFArray, PDFDict, PDFName, PDFDocument, PDFRawStream, StandardFonts, decodePDFRawStream } from 'pdf-lib';
import { readFileSync } from 'node:fs';
import fontkit from '@pdf-lib/fontkit';
import decompress from 'woff2-encoder/decompress';
import { drawWordPdfText, finalizeWordPdfFonts, wordPdfPaintText, type WordPdfFullFontGlyphs } from './word-pdf-text';
import { wordPdfDrawFontKey, type WordPdfSnapshot } from './word-pdf-model';

it.each([true, false])('preserves literal-hyphen ink and Unicode beside its aliased minus glyph (subset=%s)', async subset => {
  const bytes = new Uint8Array(await decompress(Uint8Array.from(readFileSync(
    'node_modules/@fontsource-variable/source-serif-4/files/source-serif-4-latin-wght-normal.woff2'))));
  const source = fontkit.create(bytes);
  expect(source.glyphForCodePoint(0xad).id).toBe(source.glyphForCodePoint(0x2d).id);
  const document = await PDFDocument.create(); document.registerFontkit(fontkit);
  const page = document.addPage([300, 200]), fonts = new Map(), mappings: WordPdfFullFontGlyphs = new Map();
  const glyphs: WordPdfSnapshot['pages'][number]['glyphs'] = [...'A-\u00adB'].map((text, i) => ({
    text, x: 40 + i * 10, y: 60, size: 10, family: 'Source Serif 4', color: [0, 0, 0],
    ...(text === '\u00ad' ? { literalHyphen: true as const } : {}),
  }));
  for (const glyph of glyphs) {
    const key = wordPdfDrawFontKey(glyph); if (fonts.has(key)) continue;
    fonts.set(key, await document.embedFont(bytes, { subset: subset && !glyph.literalHyphen }));
    if (!subset || glyph.literalHyphen) mappings.set(key, new Map());
  }
  drawWordPdfText(document, page, { width: 300, height: 200, glyphs }, fonts, mappings);
  await finalizeWordPdfFonts(document, fonts, mappings);
  const literal = fonts.get(wordPdfDrawFontKey(glyphs[2]))!;
  const saved = await PDFDocument.load(await document.save()), parent = saved.context.lookup(literal.ref, PDFDict);
  expect(parent.lookup(PDFName.of('Encoding'), PDFName).toString()).toBe('/WinAnsiEncoding');
  expect(parent.has(PDFName.of('ToUnicode'))).toBe(false);
  const fontProgram = parent.lookup(PDFName.of('FontDescriptor'), PDFDict).lookup(PDFName.of('FontFile2'));
  if (!(fontProgram instanceof PDFRawStream)) throw Error('Missing emitted font program');
  const program = decodePDFRawStream(fontProgram).decode();
  expect(Uint8Array.from(program)).toEqual(bytes);
  // Inspect actual emitted font outlines: mapping a blank space to U+00AD
  // would preserve text while still losing the visible hyphen.
  const painted = fontkit.create(program).glyphForCodePoint(0xad);
  expect(painted.path.toSVG()).toMatch(/[ML]/);
  expect(painted.advanceWidth).toBe(source.glyphForCodePoint(0x2d).advanceWidth);
  expect(wordPdfPaintText('\u00ad')).toBe('-');
});

it.each([
  [' A B  tail. ', 4],
  [' A\u00a0B  tail. ', 6],
] as const)('encodes every measured whitespace advance in %j across paint boundaries', async (text, objects) => {
  const document = await PDFDocument.create();
  const font = await document.embedFont(StandardFonts.Helvetica);
  const page = document.addPage([595.32, 841.92]);
  let x = 72;
  const glyphs: WordPdfSnapshot['pages'][number]['glyphs'] = [...text].map((text, index) => {
    const glyph = { text, x, y: 104, size: 10, family: 'Arial', color: [0, 0, 0] as [number, number, number],
      decoration: index >= 1 && index <= 3 ? 1 : 0 };
    x += font.widthOfTextAtSize(text, 10);
    return glyph;
  });
  drawWordPdfText(document, page, { width: 595.32, height: 841.92, glyphs }, new Map([['Arial', font]]), new Map());
  const saved = await PDFDocument.load(await document.save());
  const contents = saved.getPages()[0].node.Contents() as PDFArray;
  const operators = contents.asArray().map(ref => new TextDecoder().decode(
    decodePDFRawStream(saved.context.lookup(ref) as PDFRawStream).decode())).join('\n');
  const encoded = [...operators.matchAll(/<([\da-f]+)>/gi)].map(match => Buffer.from(match[1], 'hex').toString('ascii')).join('');
  expect(encoded).toBe(text.replace(/\u00a0/g, ' '));
  // Independent Word controls retain leading boundary whitespace in its own
  // text object; spaces remain actual encoded characters, never deleted.
  expect(operators.match(/\bBT\b/g)).toHaveLength(objects);
  expect(glyphs.map(glyph => glyph.text).join('')).toBe(text);
});
