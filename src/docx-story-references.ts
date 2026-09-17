import type JSZip from 'jszip';
import type { WordContent } from './model';
import { WORD_NS, descendants, child, val, wordXml, wElement, type readDocx } from './docx-import';
import { wordStoriesSchema } from './word-stories';
import { emptyStoryDocument } from './docx-empty-stories';
import { ensureStoryStylePart } from './docx-story-style-part';
import { clonedStoryHtml, clonedStoryParagraphKey } from './word-story-links';
import { wordStoryPartTarget } from './word-story-source';
import { wordEditSession } from './docx-edit-session';
import { resolveWordSections } from './word-section-layout';
import { liveSectionProperties } from './docx-sections';

const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const relPath = (path: string) => path.replace(/([^/]+)$/, '_rels/$1.rels');

/** Seed private cloned XML/paragraph maps before the ordinary retained editor
 * compares changes. Original parts, document identities and bytes stay intact. */
export async function prepareWordStoryClones(
  zip: JSZip,
  source: Awaited<ReturnType<typeof readDocx>>,
  current: WordContent,
) {
  const stories = current.stories && wordStoriesSchema.parse(current.stories);
  const originals = [...(source.content.stories?.parts || [])];
  const copies = (stories?.parts || []).filter((p) => !originals.some((o) => o.path === p.path));
  if (!copies.length) return false;
  if (copies.some((copy) => copy.created)) await ensureStoryStylePart(zip);
  const relationships = wordXml(await zip.file('word/_rels/document.xml.rels')!.async('string'));
  const types = wordXml(await zip.file('[Content_Types].xml')!.async('string'));
  if (
    relationships.documentElement.namespaceURI !== REL ||
    types.documentElement.namespaceURI !== CT
  )
    throw Error('Cannot add a story to invalid package relationships or content types.');
  const session = await wordEditSession(zip);
  for (const copy of copies) {
    const original = originals.find((p) => p.path === copy.copiedFrom && p.kind === copy.kind);
    if (!copy.created && (!original || !copy.copiedFrom))
      throw Error('The new header/footer has no retained source to copy.');
    if (
      zip.file(copy.path) ||
      zip.file(relPath(copy.path)) ||
      [...relationships.documentElement.children].some(
        (r) => r.getAttribute('Id') === copy.relationshipIds[0],
      ) ||
      [...types.documentElement.children].some(
        (t) => t.getAttribute('PartName') === `/${copy.path}`,
      )
    )
      throw Error('The new header/footer identity collides with a retained package part.');
    const originalDocument = copy.created
      ? await emptyStoryDocument(zip, copy.kind)
      : source.documents.get(original!.path)!;
    const document = originalDocument.cloneNode(true) as XMLDocument;
    const oldParagraphs = descendants(originalDocument, 'p'),
      paragraphs = descendants(document, 'p');
    const keys = new Map([...source.paragraphs].map(([key, paragraph]) => [paragraph, key]));
    const identity = session(copy.path);
    paragraphs.forEach((paragraph, i) => {
      const key = copy.created ? 'empty:0' : keys.get(oldParagraphs[i]);
      if (key) source.paragraphs.set(clonedStoryParagraphKey(copy.path, key), paragraph);
      if (identity) {
        paragraph.setAttributeNS(WORD_NS, 'w:rsidR', identity);
        paragraph.setAttributeNS(WORD_NS, 'w:rsidRDefault', identity);
      }
    });
    source.documents.set(copy.path, document);
    const baseline = copy.created
      ? { ...copy, created: undefined, html: source.content.stories!.emptyTemplates![copy.kind]! }
      : original!;
    source.content.stories!.parts.push({ ...copy, html: clonedStoryHtml(baseline, copy.path) });
    zip.file(copy.path, new XMLSerializer().serializeToString(document), { createFolders: false });
    const oldRelationships = !copy.created && zip.file(relPath(original!.path));
    if (oldRelationships) {
      const refs = wordXml(await oldRelationships.async('string'));
      if (refs.documentElement.namespaceURI !== REL)
        throw Error('The copied story relationships are invalid.');
      for (const ref of refs.documentElement.children) {
        if (ref.getAttribute('TargetMode') === 'External') continue;
        const target = wordStoryPartTarget(original!.path, ref.getAttribute('Target') || '');
        ref.setAttribute('Target', `/${target}`);
      }
      zip.file(relPath(copy.path), new XMLSerializer().serializeToString(refs), {
        createFolders: false,
      });
    }
    const relationship = relationships.createElementNS(REL, 'Relationship');
    relationship.setAttribute('Id', copy.relationshipIds[0]);
    relationship.setAttribute('Type', `${OFFICE}/${copy.kind}`);
    relationship.setAttribute('Target', copy.path.slice('word/'.length));
    relationships.documentElement.append(relationship);
    const type = types.createElementNS(CT, 'Override');
    type.setAttribute('PartName', `/${copy.path}`);
    type.setAttribute(
      'ContentType',
      `application/vnd.openxmlformats-officedocument.wordprocessingml.${copy.kind}+xml`,
    );
    types.documentElement.append(type);
  }
  zip.file('word/_rels/document.xml.rels', new XMLSerializer().serializeToString(relationships), {
    createFolders: false,
  });
  zip.file('[Content_Types].xml', new XMLSerializer().serializeToString(types), {
    createFolders: false,
  });
  return true;
}

/** Apply final effective references after live section movement. Removing a
 * direct reference links the slot; unrelated parts remain in the archive. */
export function writeWordStoryReferences(body: Element, current: WordContent) {
  if (!current.stories?.references?.length) return false;
  const sections = resolveWordSections(current),
    properties = liveSectionProperties(body);
  if (sections.length !== properties.length)
    throw Error('The live story sections cannot be mapped safely.');
  const known = new Set(sections.map((s) => s.id));
  // A surviving first section may deliberately have no story after deletion.
  // Its inherited override removes the old direct reference; the UI still
  // rejects a user Link-to-previous command in the first section.
  if (current.stories.references.some((r) => !known.has(r.sectionId)))
    throw Error('A header/footer link refers to an unavailable previous section.');
  let changed = false;
  for (const [index, section] of sections.entries()) {
    const propertiesForSection = properties[index];
    for (const [kind, key] of [
      ['header', 'headers'],
      ['footer', 'footers'],
    ] as const) {
      for (const slot of ['default', 'even', 'first'] as const) {
        const effective = section[key][slot];
        const wanted = effective && !effective.inherited ? effective.relationshipId : null;
        const existing = [...propertiesForSection.children].filter(
          (r) =>
            r.namespaceURI === WORD_NS &&
            r.localName === `${kind}Reference` &&
            val(r, 'type') === slot,
        );
        if (existing.length > 1) throw Error('The source has ambiguous header/footer references.');
        if (
          (!wanted && !existing.length) ||
          (wanted && existing[0]?.getAttributeNS(OFFICE, 'id') === wanted)
        )
          continue;
        existing.forEach((r) => r.remove());
        if (wanted) {
          const reference = wElement(body.ownerDocument, `${kind}Reference`, { type: slot });
          reference.setAttributeNS(OFFICE, 'r:id', wanted);
          const anchor = [...propertiesForSection.children].find((e) =>
            kind === 'header'
              ? e.localName !== 'headerReference'
              : !['headerReference', 'footerReference'].includes(e.localName),
          );
          propertiesForSection.insertBefore(reference, anchor || null);
        }
        changed = true;
      }
    }
  }
  return changed;
}
