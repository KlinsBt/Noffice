import type JSZip from 'jszip';
import { wordXml, WORD_NS } from './word-xml';

export const emptyStylesXml = `<w:styles xmlns:w="${WORD_NS}"/>`;

/** Register a genuinely absent styles part before story relationship documents
 * are loaded. Never repoint an existing relationship to another style source. */
export async function ensureStoryStylePart(zip: JSZip) {
  if (zip.file('word/styles.xml')) return;
  const relPath = 'word/_rels/document.xml.rels';
  const relText = await zip.file(relPath)?.async('string');
  const typeText = await zip.file('[Content_Types].xml')?.async('string');
  if (!relText || !typeText) throw Error('Creating this story requires valid package metadata.');
  const rels = wordXml(relText),
    types = wordXml(typeText);
  const relNS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const typeNS = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const styleType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
  if (rels.documentElement.namespaceURI !== relNS || types.documentElement.namespaceURI !== typeNS)
    throw Error('Creating this story requires valid package metadata.');
  if (
    [...rels.documentElement.children].some((r) => r.getAttribute('Type') === styleType) ||
    [...types.documentElement.children].some(
      (t) => t.getAttribute('PartName') === '/word/styles.xml',
    )
  )
    throw Error('The missing styles part has conflicting package metadata.');
  const occupied = new Set([...rels.documentElement.children].map((r) => r.getAttribute('Id')));
  let id = 'rIdNofficeStyles',
    index = 0;
  while (occupied.has(id)) id = 'rIdNofficeStyles' + ++index;
  const relationship = rels.createElementNS(relNS, 'Relationship');
  relationship.setAttribute('Id', id);
  relationship.setAttribute('Type', styleType);
  relationship.setAttribute('Target', 'styles.xml');
  rels.documentElement.append(relationship);
  const override = types.createElementNS(typeNS, 'Override');
  override.setAttribute('PartName', '/word/styles.xml');
  override.setAttribute(
    'ContentType',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  );
  types.documentElement.append(override);
  const serialize = (doc: XMLDocument) => new XMLSerializer().serializeToString(doc);
  zip.file(relPath, serialize(rels), { createFolders: false });
  zip.file('[Content_Types].xml', serialize(types), { createFolders: false });
  zip.file('word/styles.xml', emptyStylesXml, { createFolders: false });
}
