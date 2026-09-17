import type { WordContent } from './model';
import { liveSectionProperties } from './docx-sections';
import { validateWordSectionIdentities } from './word-section-breaks';
import { WORD_NS, child } from './docx-import';

/** Apply explicit live geometry after global controls, on the private export XML.
 * Header/footer distances, unrelated section properties and source parts survive. */
export function writeWordSectionGeometry(body: Element, content: WordContent) {
  const state = content.sectionState;
  if (state?.version !== 2 || !state.layouts?.length) return false;
  const ids = validateWordSectionIdentities(state, content.docxStructure);
  const parts = liveSectionProperties(body);
  if (parts.length !== ids.length) throw Error('Section geometry cannot be mapped safely.');
  for (const layout of state.layouts) {
    const section = parts[ids.indexOf(layout.sectionId)];
    if (child(section, 'sectPrChange'))
      throw Error('Editing tracked section properties is not supported yet.');
    for (const [name, values] of [
      ['pgSz', { w: layout.width, h: layout.height, orient: layout.orientation }],
      ['pgMar', layout.margins],
    ] as const) {
      let property = child(section, name);
      if (!property) {
        property = section.ownerDocument.createElementNS(WORD_NS, `w:${name}`);
        const before = ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type'];
        if (name === 'pgMar') before.push('pgSz');
        const anchor = [...section.children].find(
          (e) => e.namespaceURI === WORD_NS && !before.includes(e.localName),
        );
        section.insertBefore(property, anchor || null);
      }
      for (const [key, value] of Object.entries(values))
        property.setAttributeNS(WORD_NS, `w:${key}`, String(value));
    }
  }
  return true;
}
