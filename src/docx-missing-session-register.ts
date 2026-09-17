import type JSZip from 'jszip';
import { readWordSettings } from './docx-settings';
import { WORD_NS, wElement, wordXml } from './word-xml';

/** A first saved edit must make retained paragraph/run sessions recognizable to
 * Word on its next open. Otherwise Word can merge a later native edit into this
 * save and change its glyph advances. Register existing IDs only; do not change
 * runs or activate an orphan settings part. Sources without IDs remain separate. */
export async function registerMissingWordSessions(zip: JSZip, documents: Iterable<XMLDocument>) {
  if (await readWordSettings(zip)) return false;
  const ids = new Set<string>();
  for (const doc of documents) for (const node of doc.getElementsByTagNameNS(WORD_NS, '*')) {
    for (const attr of node.attributes) {
      if (attr.namespaceURI !== WORD_NS || !attr.localName.startsWith('rsid')) continue;
      if (!/^[\dA-F]{8}$/i.test(attr.value)) return false;
      ids.add(attr.value.toUpperCase());
      if (ids.size > 65536) return false;
    }
  }
  return createWordSessionSettings(zip, [...ids]);
}

/** Create an active settings part for an otherwise unregistered package. An
 * optional root is used only when initializing the document's first session. */
export async function createWordSessionSettings(zip: JSZip, values: readonly string[], root?: string) {
  if (await readWordSettings(zip)) return false;
  if (!values.length || values.length > 65536 || values.some(value => !/^[\dA-F]{8}$/i.test(value)))
    return false;
  const ids = new Set(values.map(value => value.toUpperCase()));
  if (root !== undefined && (!/^[\dA-F]{8}$/i.test(root) || !ids.has(root.toUpperCase()))) return false;
  const relNS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const typeNS = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const relPath = 'word/_rels/document.xml.rels';
  const typeText = await zip.file('[Content_Types].xml')?.async('string');
  if (!typeText) throw Error('Saving Word session identities requires valid package metadata.');
  const types = wordXml(typeText);
  const rels = wordXml(await zip.file(relPath)?.async('string') || `<Relationships xmlns="${relNS}"/>`);
  if (types.documentElement.namespaceURI !== typeNS || types.documentElement.localName !== 'Types' ||
      rels.documentElement.namespaceURI !== relNS || rels.documentElement.localName !== 'Relationships')
    throw Error('Saving Word session identities requires valid package metadata.');
  const occupied = new Set([
    ...Object.keys(zip.files),
    ...[...types.documentElement.children].map((node) => (node.getAttribute('PartName') || '').replace(/^\//, '')),
  ].map((path) => path.toLowerCase()));
  let path = 'word/nofficeSettings.xml', index = 0;
  while (occupied.has(path.toLowerCase())) path = `word/nofficeSettings${++index}.xml`;
  const relationshipIds = new Set([...rels.documentElement.children].map((node) => node.getAttribute('Id')));
  let id = 'rIdNofficeSettings';index = 0;
  while (relationshipIds.has(id)) id = `rIdNofficeSettings${++index}`;
  const settings = wordXml(`<w:settings xmlns:w="${WORD_NS}"/>`), register = wElement(settings, 'rsids');
  if (root !== undefined) register.append(wElement(settings, 'rsidRoot', { val: root.toUpperCase() }));
  for (const value of [...ids].sort()) register.append(wElement(settings, 'rsid', { val: value }));
  settings.documentElement.append(register);
  const relationship = rels.createElementNS(relNS, 'Relationship');
  relationship.setAttribute('Id', id);
  relationship.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings');
  relationship.setAttribute('Target', path.slice('word/'.length));
  rels.documentElement.append(relationship);
  const override = types.createElementNS(typeNS, 'Override');
  override.setAttribute('PartName', '/' + path);
  override.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  types.documentElement.append(override);
  const serialize = (doc: XMLDocument) => new XMLSerializer().serializeToString(doc);
  // Prepare all XML before mutating the disposable export package.
  const output = [[path, serialize(settings)], [relPath, serialize(rels)], ['[Content_Types].xml', serialize(types)]];
  for (const [name, text] of output) zip.file(name, text, { createFolders: false });
  return true;
}
