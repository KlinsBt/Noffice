import type JSZip from 'jszip';
import { WORD_NS, descendants, val, wElement, wordXml } from './docx-import';
import { readWordSettings } from './docx-settings';

/** Stamp eligible inserted runs with this export's save-session identity. Only
 * extend an existing native rsids register; full revision authoring is separate.
 * Preparing a session never changes the package until a run claims its ID.
 */
export async function wordEditSession(
  zip: JSZip,
): Promise<(session?: string) => string | undefined> {
  const part = await readWordSettings(zip);
  if (!part) return () => undefined;
  const settings = part.document;
  const register = descendants(settings, 'rsids')[0];
  if (!register) return () => undefined;
  let maximum = 0;
  for (const entry of Object.values(zip.files)) {
    if (entry.dir || !entry.name.endsWith('.xml')) continue;
    const text = await entry.async('string');
    if (!text.includes(WORD_NS)) continue;
    let doc: XMLDocument;
    try {
      doc = entry.name === part.path ? settings : wordXml(text);
    } catch {
      // Optional session metadata cannot make an opaque retained part editable
      // or introduce a new export rejection for that unrelated part.
      return () => undefined;
    }
    for (const node of doc.getElementsByTagNameNS(WORD_NS, '*')) {
      const ids = [...node.attributes]
        .filter((a) => a.namespaceURI === WORD_NS && a.localName.startsWith('rsid'))
        .map((a) => a.value);
      if (node.localName === 'rsid') ids.push(val(node));
      for (const id of ids) {
        if (!/^[\dA-F]{8}$/i.test(id)) return () => undefined;
        maximum = Math.max(maximum, Number.parseInt(id, 16));
      }
    }
  }
  if (maximum === 0xffffffff) return () => undefined;
  const sessions = new Map<string, string>();
  const used = new Set<number>();
  return (session = '') => {
    let current = sessions.get(session);
    if (!current) {
      if (used.size >= 0xffffffff - maximum) return undefined;
      // Persisted typing provenance must select the same native ID when the
      // unchanged document is exported again after reload. Untracked legacy
      // insertions still receive a fresh export-local identity.
      let seed = 2166136261;
      for (const character of session) seed = Math.imul(seed ^ character.charCodeAt(0), 16777619);
      const random =
        (session ? seed >>> 0 : crypto.getRandomValues(new Uint32Array(1))[0]) / 0x100000000;
      let candidate = maximum + 1 + Math.floor(random * (0xffffffff - maximum));
      while (used.has(candidate))
        candidate = candidate === 0xffffffff ? maximum + 1 : candidate + 1;
      used.add(candidate);
      current = candidate.toString(16).padStart(8, '0').toUpperCase();
      sessions.set(session, current);
      register.append(wElement(settings, 'rsid', { val: current }));
      zip.file(part.path, new XMLSerializer().serializeToString(settings), {
        createFolders: false,
      });
    }
    return current;
  };
}
