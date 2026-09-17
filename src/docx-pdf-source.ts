import JSZip from 'jszip';
import { inspectZip } from './formats';
import { docxStyles } from './docx-styles';
import { child, descendants, val, wordXml, WORD_NS } from './word-xml';

/** The HTML reader currently reduces Word's complex underline styles to a
 * boolean and omits double strike. Check the retained source before allowing
 * a PDF to silently claim that simplified appearance, including old saves.
 * Source styles remain a conservative limit even after local formatting edits. */
export async function assertWordPdfSourceDecorations(data: ArrayBuffer) {
  inspectZip(data);
  const zip = await JSZip.loadAsync(data);
  const styleParagraph = docxStyles(await zip.file('word/styles.xml')?.async('string'));
  for (const part of Object.values(zip.files)) {
    if (part.dir || !part.name.startsWith('word/') || !part.name.endsWith('.xml')) continue;
    const xml = await part.async('string');
    // Parse bounded XML before inspecting namespaces; filenames are not part identities.
    const doc = wordXml(xml), root = doc.documentElement;
    if (root.namespaceURI !== WORD_NS || !['document', 'hdr', 'ftr', 'footnotes', 'endnotes'].includes(root.localName)) continue;
    for (const paragraph of descendants(doc, 'p')) {
      styleParagraph(paragraph);
      for (const run of descendants(paragraph, 'r')) {
        const properties = child(run, 'rPr');
        const underline = properties && child(properties, 'u'), doubleStrike = properties && child(properties, 'dstrike');
        const underlineValue = val(underline);
        if ((underline && !['none', '0', 'false', 'off'].includes(underlineValue)
          && (!['', 'single'].includes(underlineValue)
            || [...underline.attributes].some(attribute => attribute.namespaceURI === WORD_NS
              && ['color', 'themeColor', 'themeTint', 'themeShade'].includes(attribute.localName)
              && !(attribute.localName === 'color' && attribute.value === 'auto'))))
          || (doubleStrike && !['0', 'false', 'off'].includes(val(doubleStrike))))
          throw Error('PDF export does not yet support the original DOCX\'s complex underline or double-strikethrough styles. Export DOCX to preserve them.');
      }
    }
  }
}
