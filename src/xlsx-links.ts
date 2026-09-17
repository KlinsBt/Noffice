import type JSZip from 'jszip';
import type { Cell } from './model';
import { address, coordinates } from './formulas';
import { sheetLinkAddress } from './sheet-links';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const all = (root: Document | Element, name: string) =>
  Array.from(root.getElementsByTagNameNS('*', name));
const xml = (text: string) => {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw Error('Invalid hyperlink XML.');
  return doc;
};
const relationPath = (path: string) => path.replace(/([^/]+)$/, '_rels/$1.rels');
function rect(ref: string) {
  const [a, b = a] = ref.split(':');
  const [r, c] = coordinates(a),
    [rr, cc] = coordinates(b);
  return [Math.min(r, rr), Math.min(c, cc), Math.max(r, rr), Math.max(c, cc)];
}
function includes(range: number[], r: number, c: number) {
  return r >= range[0] && r <= range[2] && c >= range[1] && c <= range[3];
}
function removeCell(range: number[], r: number, c: number) {
  const [top, left, bottom, right] = range;
  return [
    [top, left, r - 1, right],
    [r + 1, left, bottom, right],
    [r, left, r, c - 1],
    [r, c + 1, r, right],
  ]
    .filter(([a, b, x, y]) => a <= x && b <= y)
    .map(([a, b, x, y]) => `${address(a, b)}:${address(x, y)}`);
}
export function readWorksheetLinks(
  root: Element,
  relations: string | undefined,
  cells: Record<string, Cell>,
) {
  const targets = new Map(
    relations
      ? all(xml(relations), 'Relationship')
          .filter((e) => e.getAttribute('Type')?.endsWith('/hyperlink'))
          .map((e) => [e.getAttribute('Id'), e.getAttribute('Target') || ''])
      : [],
  );
  let scanned = 0;
  for (const node of all(root, 'hyperlink')) {
    const ref = node.getAttribute('ref') || '';
    let range: number[];
    try {
      range = rect(ref);
    } catch {
      continue;
    }
    const target = targets.get(node.getAttributeNS(R, 'id') || node.getAttribute('r:id')) || '';
    const location = node.getAttribute('location');
    const href = target + (location ? '#' + location : '');
    if (!href) continue;
    const single = range[0] === range[2] && range[1] === range[3];
    const singleRef = address(range[0], range[1]);
    if (single && range[0] < 10000 && range[1] < 256)
      cells[singleRef] ||= { value: node.getAttribute('display') || '', dataType: 'text' };
    for (const [key, cell] of single
      ? cells[singleRef]
        ? ([[singleRef, cells[singleRef]]] as [string, Cell][])
        : []
      : Object.entries(cells)) {
      if (++scanned > 5000000)
        throw Error('This worksheet has too many overlapping hyperlink ranges to import.');
      const [r, c] = coordinates(key);
      if (!includes(range, r, c)) continue;
      cell.hyperlink = href;
      cell.hyperlinkTooltip = node.getAttribute('tooltip') || undefined;
    }
  }
}
/** Patch changed links while retaining unrelated worksheet nodes and relationship IDs. */
export async function writeWorksheetLinks(
  zip: JSZip,
  path: string,
  root: Element,
  before: Record<string, Cell>,
  after: Record<string, Cell>,
) {
  const refs = [...new Set([...Object.keys(before), ...Object.keys(after)])].filter(
    (ref) =>
      before[ref]?.hyperlink !== after[ref]?.hyperlink ||
      before[ref]?.hyperlinkTooltip !== after[ref]?.hyperlinkTooltip ||
      (after[ref]?.hyperlink && before[ref]?.value !== after[ref]?.value),
  );
  if (!refs.length) return false;
  const doc = root.ownerDocument,
    relPath = relationPath(path),
    source = await zip.file(relPath)?.async('string');
  const rels = xml(source || `<Relationships xmlns="${PR}"/>`);
  let relDirty = false;
  const links =
    Array.from(root.children).find((e) => e.localName === 'hyperlinks') ||
    doc.createElementNS(root.namespaceURI, 'hyperlinks');
  if (!links.parentNode) {
    const later = [
      'printOptions',
      'pageMargins',
      'pageSetup',
      'headerFooter',
      'rowBreaks',
      'colBreaks',
      'customProperties',
      'cellWatches',
      'ignoredErrors',
      'smartTags',
      'drawing',
      'legacyDrawing',
      'legacyDrawingHF',
      'picture',
      'oleObjects',
      'controls',
      'webPublishItems',
      'tableParts',
      'extLst',
    ];
    root.insertBefore(
      links,
      Array.from(root.children).find((e) => later.includes(e.localName)) || null,
    );
  }
  for (const ref of refs) {
    const [r, c] = coordinates(ref);
    let template: Element | undefined;
    for (const node of Array.from(links.children)) {
      const range = rect(node.getAttribute('ref') || '');
      if (!includes(range, r, c)) continue;
      template ||= node.cloneNode(true) as Element;
      for (const remaining of removeCell(range, r, c)) {
        const clone = node.cloneNode(true) as Element;
        clone.setAttribute('ref', remaining);
        links.insertBefore(clone, node);
      }
      node.remove();
    }
    const next = after[ref];
    if (!next?.hyperlink) continue;
    // Unchanged legacy file destinations can retain ScreenTips; only new destinations require validation.
    const href =
      before[ref]?.hyperlink === next.hyperlink ? next.hyperlink : sheetLinkAddress(next.hyperlink);
    const node = template || doc.createElementNS(root.namespaceURI, 'hyperlink');
    node.setAttribute('ref', ref);
    if (!template || before[ref]?.hyperlink !== next.hyperlink) {
      node.removeAttributeNS(R, 'id');
      node.removeAttribute('r:id');
      node.removeAttribute('location');
      if (href.startsWith('#')) node.setAttribute('location', href.slice(1));
      else {
        let rel = all(rels, 'Relationship').find(
          (e) =>
            e.getAttribute('Type') === R + '/hyperlink' &&
            e.getAttribute('TargetMode') === 'External' &&
            e.getAttribute('Target') === href,
        );
        if (!rel) {
          const ids = new Set(all(rels, 'Relationship').map((e) => e.getAttribute('Id')));
          let n = 1;
          while (ids.has('rId' + n)) n++;
          rel = rels.createElementNS(PR, 'Relationship');
          for (const [key, value] of Object.entries({
            Id: 'rId' + n,
            Type: R + '/hyperlink',
            TargetMode: 'External',
            Target: href,
          }))
            rel.setAttribute(key, value);
          rels.documentElement.appendChild(rel);
          relDirty = true;
        }
        node.setAttributeNS(R, 'r:id', rel.getAttribute('Id')!);
      }
    }
    if (next.hyperlinkTooltip) node.setAttribute('tooltip', next.hyperlinkTooltip);
    else node.removeAttribute('tooltip');
    if (!next.value.startsWith('=') || next.dataType === 'text')
      node.setAttribute('display', next.value);
    links.appendChild(node);
  }
  if (!links.children.length) links.remove();
  // Keep other relationships, including shared or opaque references; prune only orphaned hyperlink IDs.
  const used = new Set(
    all(root, '*').flatMap((e) =>
      Array.from(e.attributes)
        .filter((a) => a.namespaceURI === R || a.name === 'r:id')
        .map((a) => a.value),
    ),
  );
  for (const rel of all(rels, 'Relationship'))
    if (rel.getAttribute('Type') === R + '/hyperlink' && !used.has(rel.getAttribute('Id')!)) {
      rel.remove();
      relDirty = true;
    }
  if (relDirty)
    zip.file(relPath, new XMLSerializer().serializeToString(rels), { createFolders: false });
  return true;
}
