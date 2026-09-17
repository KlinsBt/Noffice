import { createRequire } from 'node:module';
import { PDFDocument, PDFArray, PDFDict, PDFName, PDFNumber, PDFString, PDFHexString,
  PDFRawStream, decodePDFRawStream } from 'pdf-lib';

const require = createRequire(import.meta.url);
// Pinned pdf-lib 1.17.1 supplies PDF string/array escaping and token parsing.
const ByteStream = require('pdf-lib/cjs/core/parser/ByteStream.js').default;
const ObjectParser = require('pdf-lib/cjs/core/parser/PDFObjectParser.js').default;
const name = value => PDFName.of(value);
const string = value => value instanceof PDFString || value instanceof PDFHexString;
const decoded = stream => {
  if (!(stream instanceof PDFRawStream)) throw Error('Expected a PDF content stream');
  const bytes = decodePDFRawStream(stream).decode();
  if (bytes.length > 8_000_000) throw Error('PDF evidence stream exceeds its limit');
  return bytes;
};

function fontDecoder(font) {
  const cmap = font.lookupMaybe(name('ToUnicode'), PDFRawStream);
  if (cmap) {
    const text = Buffer.from(decoded(cmap)).toString('ascii');
    // Current full/subset browser fonts use fixed two-byte bfchar mappings.
    // Reject other encodings instead of guessing and losing characters.
    if (text.includes('beginbfrange') || !/1\s+begincodespacerange\s*<0000>\s*<ffff>\s*endcodespacerange/i.test(text))
      throw Error('Unsupported PDF evidence character map');
    const mapping = new Map();
    for (const block of text.matchAll(/(\d+)\s+beginbfchar([\s\S]*?)endbfchar/g)) {
      const entries = [...block[2].matchAll(/<([\da-f]{4})>\s*<((?:[\da-f]{4})+)>/gi)];
      if (entries.length !== Number(block[1])) throw Error('Incomplete PDF character map');
      for (const entry of entries) {
        const key = Number.parseInt(entry[1], 16);
        if (mapping.has(key)) throw Error('Duplicate PDF character mapping');
        mapping.set(key, entry[2].match(/.{4}/g).map(unit => String.fromCharCode(Number.parseInt(unit, 16))).join(''));
      }
    }
    return bytes => {
      if (bytes.length % 2) throw Error('Incomplete PDF character code');
      let text = '';
      for (let i = 0; i < bytes.length; i += 2) {
        const value = mapping.get(bytes[i] * 256 + bytes[i + 1]);
        if (value === undefined) throw Error('Unmapped PDF character');
        text += value;
      }
      return text;
    };
  }
  if (font.lookupMaybe(name('Encoding'), PDFName)?.decodeText() !== 'WinAnsiEncoding')
    throw Error('Unsupported PDF evidence font encoding');
  return bytes => new TextDecoder('windows-1252', { fatal: true }).decode(bytes);
}

/** Read encoded text, including isolated invisible spaces omitted by PDFium's
 * extracted character list. This bounded evidence reader does not reconstruct
 * gaps, infer text order, measure glyphs, or certify arbitrary PDF documents. */
export async function encodedPdfTextObjects(input) {
  if (input.length > 25_000_000) throw Error('PDF evidence exceeds its limit');
  const document = await PDFDocument.load(input);
  if (document.getPageCount() > 1000) throw Error('Too many PDF evidence pages');
  const result = [];
  let characters = 0, tokens = 0;
  for (const page of document.getPages()) {
    const resources = page.node.Resources().lookup(name('Font'), PDFDict);
    const decoders = new Map();
    const contents = page.node.Contents();
    const streams = contents instanceof PDFArray ? contents.asArray().map(ref => document.context.lookup(ref)) : [contents];
    const objects = [], operands = [];
    let active, font, size;
    for (const stream of streams) {
      const bytes = ByteStream.of(decoded(stream)), parser = ObjectParser.forByteStream(bytes, document.context);
      while (!bytes.done()) {
        parser.skipWhitespaceAndComments();
        if (bytes.done()) break;
        if (++tokens > 100_000) throw Error('Too many PDF evidence tokens');
        const char = String.fromCharCode(bytes.peek());
        if (/[\d.+\-/<\[(]/.test(char)) { operands.push(parser.parseObject()); continue; }
        let operator = '';
        while (!bytes.done() && !/[\s\x00<>\[\]()/%]/.test(String.fromCharCode(bytes.peek())))
          operator += String.fromCharCode(bytes.next());
        if (!operator) throw Error('Invalid PDF evidence operator');
        if (['Do', 'BI', 'cm', 'Td', 'TD', "'", '"'].includes(operator))
          throw Error('Unsupported transformed or indirect PDF text');
        if (operator === 'BT') {
          if (active) throw Error('Nested PDF text object');
          active = { text: '', x: 0, y: 0, positioned: false };
        } else if (operator === 'ET') {
          if (!active) throw Error('Unmatched PDF text object');
          delete active.positioned;
          objects.push(active); active = undefined;
        } else if (operator === 'T*') {
          if (!active) throw Error('Line movement outside PDF text');
          // pdf-lib emits a final line move before ET. No text may be read
          // from that implicit position without another supported matrix.
          active.positioned = false;
        } else if (operator === 'Tf') {
          if (operands.length !== 2 || !(operands[0] instanceof PDFName) || !(operands[1] instanceof PDFNumber)) throw Error('Invalid PDF font');
          font = operands[0].decodeText(); size = operands[1].asNumber();
          if (!decoders.has(font)) decoders.set(font, fontDecoder(resources.lookup(name(font), PDFDict)));
        } else if (operator === 'Tm') {
          if (!active || active.text || operands.length !== 6 || operands.some(value => !(value instanceof PDFNumber)))
            throw Error('Unsupported PDF text matrix');
          const [a, b, c, d, x, y] = operands.map(value => value.asNumber());
          if (a !== 1 || b !== 0 || c !== 0 || d !== 1) throw Error('Unsupported PDF text transform');
          active.x = x; active.y = y; active.positioned = true;
        } else if (operator === 'Tj' || operator === 'TJ') {
          if (!active?.positioned || !decoders.has(font)) throw Error('Unsupported PDF text position or font');
          if (operands.length !== 1 || (operator === 'TJ' && !(operands[0] instanceof PDFArray)))
            throw Error('Invalid PDF text operands');
          const values = operator === 'TJ' && operands[0] instanceof PDFArray ? operands[0].asArray() : operands;
          for (const value of values) {
            if (value instanceof PDFNumber && operator === 'TJ') continue;
            if (!string(value)) throw Error('Invalid encoded PDF text');
            const text = decoders.get(font)(value.asBytes());
            if ((characters += text.length) > 50_000) throw Error('Too much PDF evidence text');
            active.text += text;
          }
          active.font = font; active.size = size;
        }
        operands.length = 0;
      }
    }
    if (active || operands.length) throw Error('Unterminated PDF text content');
    result.push(objects);
  }
  return result;
}
