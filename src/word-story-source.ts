import type JSZip from 'jszip';
import type { DocxStructure } from './docx-sections';
import { wordXml } from './word-xml';
import { readWordSettings } from './docx-settings';
import { wordStoryPartTarget } from './word-package-path';
export { wordStoryPartTarget } from './word-package-path';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export interface WordStorySource {
  path: string;
  kind: 'header' | 'footer';
  relationshipIds: string[];
  document: XMLDocument;
}

/** Read only referenced stories. Shared references retain one source part;
 * unreachable package contents stay in the original archive, not the model. */
export async function readWordStorySources(zip: JSZip, structure: DocxStructure) {
  const wanted = new Map<string, 'header' | 'footer'>();
  for (const section of structure.sections)
    for (const [key, kind] of [
      ['headers', 'header'],
      ['footers', 'footer'],
    ] as const)
      for (const reference of section[key]) {
        if (!reference.relationshipId) throw Error('A Word story reference has no identity.');
        const prior = wanted.get(reference.relationshipId);
        if (prior && prior !== kind) throw Error('A Word story reference has conflicting types.');
        wanted.set(reference.relationshipId, kind);
      }
  let evenAndOddHeaders = false;
  const settings = await readWordSettings(zip);
  if (settings) {
    const root = settings.document.documentElement;
    if (root.namespaceURI !== W || root.localName !== 'settings')
      throw Error('The Word document settings are invalid.');
    const flag = [...root.children].find(
      (c) => c.namespaceURI === W && c.localName === 'evenAndOddHeaders',
    );
    evenAndOddHeaders =
      !!flag && !['0', 'false', 'off'].includes(flag.getAttributeNS(W, 'val') || '');
  }
  if (!wanted.size) return { evenAndOddHeaders, parts: [] as WordStorySource[] };
  if (wanted.size > 600) throw Error('The document has too many header/footer references.');
  const xml = await zip.file('word/_rels/document.xml.rels')?.async('string');
  if (!xml) throw Error('The header/footer relationships are missing.');
  const root = wordXml(xml).documentElement;
  if (root.namespaceURI !== REL || root.localName !== 'Relationships')
    throw Error('The header/footer relationships are invalid.');
  const byPath = new Map<string, WordStorySource>(),
    seen = new Set<string>();
  let totalCharacters = 0;
  for (const relationship of root.children) {
    const id = relationship.getAttribute('Id') || '';
    const kind = wanted.get(id);
    if (!kind) continue;
    if (seen.has(id)) throw Error('A header/footer relationship is ambiguous.');
    seen.add(id);
    if (
      relationship.namespaceURI !== REL ||
      relationship.localName !== 'Relationship' ||
      ![null, '', 'Internal'].includes(relationship.getAttribute('TargetMode')) ||
      relationship.getAttribute('Type') !== `${OFFICE}/${kind}`
    )
      throw Error('Invalid internal header/footer relationship.');
    const path = wordStoryPartTarget(
      'word/document.xml',
      relationship.getAttribute('Target') || '',
    );
    const shared = byPath.get(path);
    if (shared) {
      if (shared.kind !== kind) throw Error('A Word story part has conflicting types.');
      shared.relationshipIds.push(id);
      continue;
    }
    const part = await zip.file(path)?.async('string');
    if (!part) throw Error('A referenced header/footer part is missing.');
    totalCharacters += part.length;
    if (part.length > 1000000 || totalCharacters > 10000000)
      throw Error('The header/footer contents exceed the supported size.');
    const document = wordXml(part);
    if (
      document.documentElement.namespaceURI !== W ||
      document.documentElement.localName !== (kind === 'header' ? 'hdr' : 'ftr')
    )
      throw Error('A referenced header/footer part has the wrong document type.');
    byPath.set(path, { path, kind, relationshipIds: [id], document });
  }
  if (seen.size !== wanted.size) throw Error('A referenced header/footer relationship is missing.');
  return {
    evenAndOddHeaders,
    parts: [...byPath.values()].sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0)),
  };
}
