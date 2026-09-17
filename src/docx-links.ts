import type JSZip from 'jszip';
import type { JSONContent } from '@tiptap/core';
import { child, wElement, wordXml } from './docx-import';

const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
type Marks = NonNullable<JSONContent['marks']>;
const href = (marks: Marks) =>
  marks.find((mark) => mark.type === 'link')?.attrs?.href as string | undefined;
export type DocxLinks = Awaited<ReturnType<typeof docxLinks>>;

/** Relationship IDs are local to a document part, including footnotes and endnotes. */
export async function docxLinks(zip: JSZip, documents: Map<string, Document>) {
  const parts = new Map<Document, { path: string; relationships: Document; dirty: boolean }>();
  for (const [path, doc] of documents) {
    const split = path.lastIndexOf('/');
    const relPath = `${path.slice(0, split + 1)}_rels/${path.slice(split + 1)}.rels`;
    const stored = zip.file(relPath);
    parts.set(doc, {
      path: relPath,
      relationships: wordXml(
        stored ? await stored.async('string') : `<Relationships xmlns="${REL}"/>`,
      ),
      dirty: false,
    });
  }
  function relationship(doc: Document, target: string) {
    if (
      target.length > 8192 ||
      !/^(https?:\/\/|mailto:)/i.test(target) ||
      /[\u0000-\u0020]/.test(target)
    )
      throw Error('Use a valid http://, https:// or mailto: hyperlink address.');
    const part = parts.get(doc);
    if (!part) throw Error('The hyperlink document part is missing.');
    const records = Array.from(part.relationships.documentElement.children);
    const match = records.find(
      (e) =>
        e.getAttribute('Type') === `${OFFICE}/hyperlink` &&
        e.getAttribute('TargetMode') === 'External' &&
        e.getAttribute('Target') === target,
    );
    if (match) return match.getAttribute('Id')!;
    const ids = new Set(records.map((e) => e.getAttribute('Id')));
    let id = 1;
    while (ids.has(`rIdNofficeLink${id}`)) id++;
    const entry = part.relationships.createElementNS(REL, 'Relationship');
    entry.setAttribute('Id', `rIdNofficeLink${id}`);
    entry.setAttribute('Type', `${OFFICE}/hyperlink`);
    entry.setAttribute('TargetMode', 'External');
    entry.setAttribute('Target', target);
    part.relationships.documentElement.appendChild(entry);
    part.dirty = true;
    return `rIdNofficeLink${id}`;
  }
  const identities = new WeakMap<
    Element,
    { template: Element | undefined; target: string | undefined }
  >();
  function append(
    output: Element[],
    run: Element,
    before: Marks,
    after: Marks,
    template?: Element,
  ) {
    const target = href(after),
      previous = href(before);
    const source =
      template?.parentElement?.localName === 'hyperlink' ? template.parentElement : undefined;
    const retain = source && target === previous;
    if (!target && !retain) {
      const style = child(child(run, 'rPr') || run, 'rStyle');
      if (previous && style?.getAttributeNS(run.namespaceURI, 'val') === 'Hyperlink')
        style.remove();
      output.push(run);
      return;
    }
    if (target && !previous) {
      let properties = child(run, 'rPr');
      if (!properties) {
        properties = wElement(run.ownerDocument, 'rPr');
        run.prepend(properties);
      }
      if (!child(properties, 'rStyle'))
        properties.prepend(wElement(run.ownerDocument, 'rStyle', { val: 'Hyperlink' }));
    }
    const last = output.at(-1),
      identity = last && identities.get(last);
    if (last && identity && identity.template === source && identity.target === target) {
      last.appendChild(run);
      return;
    }
    const link = retain
      ? (source.cloneNode(false) as Element)
      : wElement(run.ownerDocument, 'hyperlink');
    if (!retain && target) {
      if (target.startsWith('#')) {
        const anchor = target.slice(1);
        if (!anchor || anchor.length > 255 || /[\u0000-\u0020]/.test(anchor))
          throw Error('Invalid document bookmark link.');
        link.setAttributeNS(link.namespaceURI, 'w:anchor', anchor);
      } else link.setAttributeNS(OFFICE, 'r:id', relationship(run.ownerDocument, target));
    }
    link.appendChild(run);
    output.push(link);
    identities.set(link, { template: source, target });
  }
  function save() {
    for (const part of parts.values())
      if (part.dirty)
        zip.file(part.path, new XMLSerializer().serializeToString(part.relationships), {
          createFolders: false,
        });
  }
  return { append, save };
}
