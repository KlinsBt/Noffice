import { assertXmlComplexity } from './office-preservation';

export const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const descendants = (node: Document | Element, name: string) => [
  ...node.getElementsByTagNameNS(WORD_NS, name),
];
export const child = (el: Element, name: string) =>
  [...el.children].find((e) => e.namespaceURI === WORD_NS && e.localName === name);
export const val = (el: Element | undefined, attr = 'val') =>
  el?.getAttributeNS(WORD_NS, attr) ?? '';
export function wordXml(text: string) {
  assertXmlComplexity(text);
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length)
    throw new Error('The Word document contains invalid XML.');
  return doc;
}
export function wElement(doc: Document, name: string, attrs: Record<string, string> = {}) {
  const el = doc.createElementNS(WORD_NS, `w:${name}`);
  for (const [key, value] of Object.entries(attrs)) el.setAttributeNS(WORD_NS, `w:${key}`, value);
  return el;
}

/** Word writes preservation only when XML whitespace touches a text edge.
 * Adding it to ordinary text can change native glyph placement on first open,
 * even though the text and font properties are identical. */
export function writeWordText(text: Element, value: string) {
  const xml = 'http://www.w3.org/XML/1998/namespace';
  text.textContent = value;
  if (/^[ \t\r\n]|[ \t\r\n]$/.test(value))
    text.setAttributeNS(xml, 'xml:space', 'preserve');
  else text.removeAttributeNS(xml, 'space');
}
