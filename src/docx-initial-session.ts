import type JSZip from 'jszip';
import { readWordSettings } from './docx-settings';
import { wordEditSession } from './docx-edit-session';
import { createWordSessionSettings } from './docx-missing-session-register';
import { WORD_NS, descendants, wElement, wordXml } from './word-xml';

/** Establish the first saved session of an unregistered flat body. Loaded text
 * and its initial edits share one identity; subsequent typing sessions do not.
 * This operates on the disposable export package after an actual body edit. */
export async function initializeWordSourceSession(
  zip: JSZip,
  document: XMLDocument,
  initialSession: string | undefined,
) {
  if (!initialSession || !/^[a-f\d]{32}$/.test(initialSession)) return undefined;
  const part = await readWordSettings(zip);
  if (part && descendants(part.document, 'rsids').length) return undefined;
  const paragraphs = descendants(document, 'p');
  if (!paragraphs.length || paragraphs.some((paragraph) =>
    paragraph.parentElement?.namespaceURI !== WORD_NS || paragraph.parentElement.localName !== 'body'))
    return undefined;
  // Do not reinterpret an incomplete register or identities in opaque stories.
  for (const entry of Object.values(zip.files)) {
    if (entry.dir || !entry.name.endsWith('.xml')) continue;
    const text = await entry.async('string');
    if (!text.includes(WORD_NS)) continue;
    let xml: XMLDocument;
    try { xml = wordXml(text); } catch { return undefined; }
    for (const node of xml.getElementsByTagNameNS(WORD_NS, '*')) {
      if (node.localName.startsWith('rsid') || [...node.attributes].some((attribute) =>
        attribute.namespaceURI === WORD_NS && attribute.localName.startsWith('rsid')))
        return undefined;
    }
  }
  let seed = 2166136261;
  for (const character of initialSession) seed = Math.imul(seed ^ character.charCodeAt(0), 16777619);
  // A persisted origin reproduces its identity across exports without giving
  // unrelated documents the same root. Leave room for later save sessions.
  const id = (((seed >>> 0) % 0x00fffffe) + 1).toString(16).padStart(8, '0').toUpperCase();
  if (part) {
    const settings = part.document;
    const register = wElement(settings, 'rsids');
    register.append(wElement(settings, 'rsidRoot', { val: id }), wElement(settings, 'rsid', { val: id }));
    settings.documentElement.append(register);
    zip.file(part.path, new XMLSerializer().serializeToString(settings), { createFolders: false });
  } else if (!await createWordSessionSettings(zip, [id], id)) return undefined;
  const laterSession = await wordEditSession(zip);
  for (const paragraph of paragraphs) {
    paragraph.setAttributeNS(WORD_NS, 'w:rsidR', id);
    paragraph.setAttributeNS(WORD_NS, 'w:rsidRDefault', id);
  }
  return (session?: string) => !session || session === initialSession ? id : laterSession(session);
}
