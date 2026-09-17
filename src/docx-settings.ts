import type JSZip from 'jszip';
import { wordStoryPartTarget } from './word-package-path';
import { WORD_NS, wordXml } from './word-xml';

const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const SETTINGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';

/** Settings belong to the main part's relationship, not a conventional filename.
 * An unreferenced settings.xml is opaque retained content, never active settings. */
export async function readWordSettings(zip: JSZip) {
  const relationships = await zip.file('word/_rels/document.xml.rels')?.async('string');
  if (!relationships) return undefined;
  const root = wordXml(relationships).documentElement;
  if (root.namespaceURI !== REL || root.localName !== 'Relationships')
    throw Error('The Word document relationships are invalid.');
  const matches = [...root.children].filter((node) =>
    node.namespaceURI === REL && node.localName === 'Relationship' &&
    node.getAttribute('Type') === SETTINGS);
  if (!matches.length) return undefined;
  if (matches.length !== 1) throw Error('The Word document settings relationship is ambiguous.');
  const relationship = matches[0], id = relationship.getAttribute('Id');
  if (!id || [...root.children].filter((node) => node.getAttribute('Id') === id).length !== 1 ||
      ![null, '', 'Internal'].includes(relationship.getAttribute('TargetMode')))
    throw Error('The Word document settings require a unique internal relationship.');
  const path = wordStoryPartTarget('word/document.xml', relationship.getAttribute('Target') || '');
  const text = await zip.file(path)?.async('string');
  if (text === undefined) throw Error('The referenced Word document settings are missing.');
  const document = wordXml(text);
  if (document.documentElement.namespaceURI !== WORD_NS || document.documentElement.localName !== 'settings')
    throw Error('The Word document settings are invalid.');
  return { path, document };
}
