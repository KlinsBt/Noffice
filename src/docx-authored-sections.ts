import type JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const child = (node: Element, name: string) =>
  [...node.children].find((e) => e.namespaceURI === W && e.localName === name);

/** The fresh packer emits a separate empty paragraph for each section's
 * properties. Our model already contains the ending paragraph; attach sectPr
 * to that paragraph so exporting does not introduce extra paragraph marks. */
export async function attachAuthoredSectionProperties(zip: JSZip, boundaries: number[]) {
  const part = zip.file('word/document.xml');
  if (!part) throw Error('The authored Word document has no body part.');
  const xml = new DOMParser().parseFromString(await part.async('string'), 'application/xml');
  const body = child(xml.documentElement, 'body');
  if (xml.getElementsByTagName('parsererror').length || !body)
    throw Error('Invalid authored Word body.');
  const markers = [...body.children].filter(
    (p) =>
      p.namespaceURI === W &&
      p.localName === 'p' &&
      child(p, 'pPr') &&
      child(child(p, 'pPr')!, 'sectPr'),
  );
  if (markers.length !== boundaries.length)
    throw Error('The authored section count changed during export.');
  const targets: Element[] = [];
  for (const marker of markers) {
    const properties = child(marker, 'pPr')!,
      section = child(properties, 'sectPr')!;
    const previous = marker.previousElementSibling;
    if (
      marker.children.length !== 1 ||
      properties.children.length !== 1 ||
      !previous ||
      previous.namespaceURI !== W ||
      previous.localName !== 'p'
    )
      throw Error('The authored section has no safe ending paragraph.');
    let pPr = child(previous, 'pPr');
    if (!pPr) {
      pPr = xml.createElementNS(W, 'w:pPr');
      previous.prepend(pPr);
    }
    if (child(pPr, 'sectPr')) throw Error('Duplicate authored section properties.');
    pPr.insertBefore(section, child(pPr, 'pPrChange') || null);
    marker.remove();
    targets.push(previous);
  }
  const paragraphs = [...body.getElementsByTagNameNS(W, 'p')];
  if (targets.some((p, i) => paragraphs.indexOf(p) !== boundaries[i]))
    throw Error('The authored section paragraph order changed during export.');
  zip.file('word/document.xml', new XMLSerializer().serializeToString(xml));
}
