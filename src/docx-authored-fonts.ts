import type JSZip from 'jszip';
import { child, wElement, wordXml, WORD_NS } from './word-xml';

/** Fresh files already embed the bundled font. Preserve that embedding on the
 * next Word save, rather than losing it and substituting a local font. */
export async function retainAuthoredFontEmbedding(zip: JSZip) {
  const part = zip.file('word/settings.xml');
  if (!part) throw Error('The authored Word document has no settings part.');
  const doc = wordXml(await part.async('string'));
  const settings = doc.documentElement;
  if (settings.namespaceURI !== WORD_NS || settings.localName !== 'settings')
    throw Error('Invalid authored Word settings.');
  const existing = child(settings, 'embedTrueTypeFonts');
  if (existing) existing.setAttributeNS(WORD_NS, 'w:val', 'true');
  else {
    // CT_Settings places font embedding after the view/print flags and before
    // the remaining document settings. Keep the generated XML schema order.
    const earlier = new Set([
      'writeProtection',
      'view',
      'zoom',
      'removePersonalInformation',
      'removeDateAndTime',
      'doNotDisplayPageBoundaries',
      'displayBackgroundShape',
      'printPostScriptOverText',
      'printFractionalCharacterWidth',
      'printFormsData',
    ]);
    settings.insertBefore(
      wElement(doc, 'embedTrueTypeFonts'),
      [...settings.children].find((element) => !earlier.has(element.localName)) || null,
    );
  }
  zip.file(part.name, new XMLSerializer().serializeToString(doc), { createFolders: false });
}
