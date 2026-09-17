import type JSZip from 'jszip';
import { child, descendants, wElement, wordXml } from './word-xml';
import { wordStoryPartTarget, type WordStorySource } from './word-story-source';
import { assertXmlComplexity } from './office-preservation';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
export type WordStoryBoundary = { start: string; end: string };

/** Combine annotated reading copies for one bounded Mammoth conversion. These
 * temporary body blocks never become body content or retained package parts. */
export async function appendWordStoryReading(
  zip: JSZip,
  copies: Map<string, XMLDocument>,
  stories: WordStorySource[],
  token: string,
): Promise<WordStoryBoundary[]> {
  if (!stories.length) return [];
  const main = copies.get('word/document.xml')!;
  const body = descendants(main, 'body')[0];
  const finalSection = child(body, 'sectPr');
  finalSection?.remove();
  const boundaries: WordStoryBoundary[] = [];
  let mainRelationships: XMLDocument | undefined;
  let nextRelationship = 0;
  const appendBoundary = (text: string) => {
    const p = wElement(main, 'p'),
      run = wElement(main, 'r'),
      label = wElement(main, 't');
    label.textContent = text;
    run.append(label);
    p.append(run);
    body.append(p);
  };
  for (const [index, story] of stories.entries()) {
    const annotated = copies.get(story.path)!;
    const fragment = main.createDocumentFragment();
    for (const node of annotated.documentElement.childNodes)
      fragment.append(main.importNode(node, true));
    const uses = [...fragment.querySelectorAll('*')].flatMap((element) =>
      [...element.attributes].filter(
        (a) => a.namespaceURI === R && ['id', 'embed', 'link'].includes(a.localName),
      ),
    );
    if (uses.length) {
      const slash = story.path.lastIndexOf('/');
      const path = `${story.path.slice(0, slash + 1)}_rels/${story.path.slice(slash + 1)}.rels`;
      const xml = await zip.file(path)?.async('string');
      if (!xml) throw Error('A header/footer content relationship is missing.');
      const relationships = wordXml(xml).documentElement;
      if (relationships.namespaceURI !== REL || relationships.localName !== 'Relationships')
        throw Error('Invalid header/footer content relationships.');
      if (!mainRelationships) {
        const xml = await zip.file('word/_rels/document.xml.rels')?.async('string');
        mainRelationships = wordXml(xml || `<Relationships xmlns="${REL}"/>`);
      }
      const occupied = new Set(
        [...mainRelationships.documentElement.children].map((r) => r.getAttribute('Id')),
      );
      const mapped = new Map<string, string>();
      for (const attribute of uses) {
        const old = attribute.value;
        let id = mapped.get(old);
        if (!id) {
          const matches = [...relationships.children].filter(
            (r) =>
              r.namespaceURI === REL &&
              r.localName === 'Relationship' &&
              r.getAttribute('Id') === old,
          );
          if (matches.length !== 1)
            throw Error('A header/footer content relationship is missing or ambiguous.');
          const record = mainRelationships.importNode(matches[0], true) as Element;
          const mode = record.getAttribute('TargetMode');
          if (![null, '', 'Internal', 'External'].includes(mode))
            throw Error('Invalid header/footer content relationship mode.');
          const target = record.getAttribute('Target');
          if (!target) throw Error('A header/footer content target is missing.');
          if (mode !== 'External') {
            const resolved = wordStoryPartTarget(story.path, target);
            if (!zip.file(resolved)) throw Error('A header/footer content part is missing.');
            record.setAttribute(
              'Target',
              resolved.startsWith('word/') ? resolved.slice(5) : `../${resolved}`,
            );
          }
          do {
            id = `${token}REL${nextRelationship++}`;
          } while (occupied.has(id));
          occupied.add(id);
          mapped.set(old, id);
          record.setAttribute('Id', id);
          mainRelationships.documentElement.append(record);
        }
        attribute.value = id;
      }
    }
    const boundary = { start: `${token}STORY${index}BEGIN`, end: `${token}STORY${index}END` };
    boundaries.push(boundary);
    appendBoundary(boundary.start);
    body.append(fragment);
    appendBoundary(boundary.end);
  }
  if (finalSection) body.append(finalSection);
  const serialized = new XMLSerializer().serializeToString(main);
  assertXmlComplexity(serialized);
  zip.file('word/document.xml', serialized, { createFolders: false });
  if (mainRelationships)
    zip.file(
      'word/_rels/document.xml.rels',
      new XMLSerializer().serializeToString(mainRelationships),
      { createFolders: false },
    );
  return boundaries;
}

/** Extract stories before assigning fallback paragraph identities, keeping
 * legacy body/footnote identities independent of how many stories exist. */
export function splitWordStoryReading(html: Document, boundaries: WordStoryBoundary[]) {
  const output = boundaries.map(() => html.createElement('div'));
  const markers = new Map<string, { index: number; start: boolean }>(
    boundaries.flatMap(
      (b, index) =>
        [
          [b.start, { index, start: true }],
          [b.end, { index, start: false }],
        ] as const,
    ),
  );
  let active = -1,
    completed = 0;
  for (const node of [...html.body.childNodes]) {
    const marker = markers.get(node.textContent || '');
    if (marker) {
      if (
        node.nodeType !== Node.ELEMENT_NODE ||
        (marker.start ? active !== -1 || marker.index !== completed : active !== marker.index)
      )
        throw Error('Header/footer conversion boundaries are ambiguous.');
      active = marker.start ? marker.index : -1;
      if (!marker.start) completed++;
      node.remove();
    } else if (active >= 0) output[active].append(node);
  }
  if (active !== -1 || completed !== boundaries.length)
    throw Error('Header/footer conversion did not preserve story boundaries.');
  return output;
}
