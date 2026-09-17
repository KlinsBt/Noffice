import { WORD_NS, descendants, val } from './word-xml';

/** Only the collapsed, unique Word navigation bookmark is movable here.
 * Named ranges, noncollapsed ranges, columns and opaque containers retain the
 * preservation adapter's existing rejection. The offset is semantic run text. */
export function collapsedWordNavigationAt(paragraph: Element, offset: number): Element[] {
  let position = 0;
  for (const element of [...paragraph.children]) {
    if (element.namespaceURI !== WORD_NS) return [];
    if (element.localName === 'pPr') continue;
    if (element.localName === 'r') {
      for (const runChild of [...element.children]) {
        if (runChild.namespaceURI !== WORD_NS) return [];
        if (runChild.localName === 't') position += (runChild.textContent || '').length;
        else if (!['rPr', 'lastRenderedPageBreak'].includes(runChild.localName)) return [];
      }
      continue;
    }
    if (position !== offset || element.localName !== 'bookmarkStart' || val(element, 'name') !== '_GoBack') return [];
    const end = element.nextElementSibling, id = val(element, 'id');
    const plain = (node: Element, names: string[]) => !node.childNodes.length && [...node.attributes].every(a =>
      a.namespaceURI === 'http://www.w3.org/2000/xmlns/' ||
      (a.namespaceURI === WORD_NS && names.includes(a.localName)));
    if (!/^\d+$/.test(id) || !plain(element, ['id','name']) || !end ||
      end.namespaceURI !== WORD_NS || end.localName !== 'bookmarkEnd' || val(end,'id') !== id || !plain(end,['id'])) return [];
    const starts = descendants(paragraph.ownerDocument, 'bookmarkStart');
    if (starts.filter(e => val(e,'id') === id).length !== 1 ||
      starts.filter(e => val(e,'name') === '_GoBack').length !== 1 ||
      descendants(paragraph.ownerDocument, 'bookmarkEnd').filter(e => val(e,'id') === id).length !== 1) return [];
    return [element,end];
  }
  return [];
}
