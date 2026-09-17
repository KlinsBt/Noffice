import { writeWordText } from './word-xml';
const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml = 'http://www.w3.org/XML/1998/namespace';

function signature(run: Element) {
  if (run.namespaceURI !== word || run.localName !== 'r' || run.attributes.length) return null;
  const nodes = [...run.childNodes];
  if (nodes.length !== 2 || nodes.some((n) => n.nodeType !== 1)) return null;
  const [properties, text] = nodes as Element[];
  if (
    properties.namespaceURI !== word ||
    properties.localName !== 'rPr' ||
    properties.attributes.length ||
    text.namespaceURI !== word ||
    text.localName !== 't' ||
    [...text.childNodes].some((n) => n.nodeType !== 3) ||
    [...text.attributes].some(
      (a) => a.namespaceURI !== xml || a.localName !== 'space' || a.value !== 'preserve',
    )
  )
    return null;
  // Joining must not turn non-preserved edge whitespace into interior text.
  if (!text.hasAttributeNS(xml, 'space') && text.textContent !== text.textContent?.trim())
    return null;
  const names = new Set<string>();
  for (const node of properties.childNodes) {
    if (node.nodeType !== 1) return null;
    const property = node as Element;
    if (
      property.namespaceURI !== word ||
      !['rFonts', 'sz'].includes(property.localName) ||
      names.has(property.localName) ||
      property.childNodes.length
    )
      return null;
    names.add(property.localName);
    const allowed = property.localName === 'rFonts' ? ['ascii', 'hAnsi'] : ['val'];
    if (
      [...property.attributes].some(
        (a) => a.namespaceURI !== word || !allowed.includes(a.localName),
      )
    )
      return null;
  }
  if (!names.has('rFonts')) return null;
  return { key: new XMLSerializer().serializeToString(properties), text };
}

export const isPlainWordRun = (run: Element) => signature(run) !== null;

/** A saved body part must also normalize redundant whitespace metadata on
 * retained text. Word's first-open glyph positions can differ when ordinary
 * text still carries xml:space="preserve". Keep significant whitespace and
 * opaque/nested content intact; unchanged package parts never call this. */
export function normalizeWordTextSpace(document: XMLDocument) {
  let changed = 0;
  for (const paragraph of document.getElementsByTagNameNS(word, 'p')) {
    if (paragraph.parentElement?.namespaceURI !== word ||
        paragraph.parentElement.localName !== 'body') continue;
    for (const run of paragraph.children) {
      if (run.namespaceURI !== word || run.localName !== 'r' ||
          [...run.children].some((child) => child.namespaceURI !== word ||
            !['rPr', 't'].includes(child.localName))) continue;
      for (const text of run.children) {
        if (text.localName !== 't' ||
            [...text.childNodes].some((child) => child.nodeType !== 3) ||
            text.attributes.length !== 1 ||
            text.getAttributeNS(xml, 'space') !== 'preserve') continue;
        const value = text.textContent || '';
        if (!value || /^[ ]|[ ]$|[\t\r\n]/.test(value)) continue;
        text.removeAttributeNS(xml, 'space');
        changed++;
      }
    }
  }
  return changed;
}

/** Native Save coalesces the measured adjacent plain font runs. Keep opaque,
 * revision-bearing, field, bookmark and drawing boundaries untouched. This is
 * applied only to an already edited main part, never to the preserved original.
 */
export function coalescePlainWordRuns(
  document: XMLDocument,
  insertedRuns: ReadonlySet<Element> = new Set(),
) {
  let merged = 0;
  for (const paragraph of document.getElementsByTagNameNS(word, 'p')) {
    // This measured contract owns body paragraphs, not opaque drawing text,
    // table cells or nested revision/content-control containers.
    if (
      paragraph.parentElement?.namespaceURI !== word ||
      paragraph.parentElement.localName !== 'body'
    )
      continue;
    let previous: ReturnType<typeof signature> = null;
    for (const node of [...paragraph.childNodes]) {
      const current =
        node.nodeType === 1 && !insertedRuns.has(node as Element)
          ? signature(node as Element)
          : null;
      if (current && previous?.key === current.key) {
        writeWordText(previous.text, (previous.text.textContent ?? '') + (current.text.textContent ?? ''));
        paragraph.removeChild(node);
        merged++;
      } else previous = current;
    }
  }
  return merged;
}
