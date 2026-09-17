import { validateWordSectionIdentities, type WordSectionState } from './word-section-breaks';
import type { DocxStructure } from './docx-sections';
import { liveSectionProperties } from './docx-sections';
import { WORD_NS, child, val } from './docx-import';
import { materializeWordFinalSection } from './word-section-defaults';

/** Work on the exporter's private XML copy only. Return a commit after paragraphs are patched.
 * Surviving sections retain their effective header/footer slots even if their predecessor vanished.
 */
export function prepareLiveSections(body: Element, source: DocxStructure, live: WordSectionState) {
  validateWordSectionIdentities(live, source);
  const materialized = materializeWordFinalSection(body, source.compatibility?.mode);
  const properties = liveSectionProperties(body);
  const slots = new Map<string, Element>();
  const effective = new Map<string, Map<string, Element>>();
  const parts = new Map<string, Element>();
  const oldParents = new Map<string, Element>();
  const nonFinal = source.sections.filter((s) => s.endingParagraph !== null);
  if (nonFinal.length !== properties.filter((p) => p.parentElement !== body).length)
    throw Error('The source section properties cannot be mapped safely.');
  for (let i = 0; i < source.sections.length; i++) {
    const section = source.sections[i],
      part = properties[i];
    if (!part) {
      if (nonFinal.length)
        throw Error('Editing breaks before implicit final section settings is not supported yet.');
      continue;
    }
    for (const ref of [...part.children]) {
      if (
        ref.namespaceURI === WORD_NS &&
        ['headerReference', 'footerReference'].includes(ref.localName)
      )
        slots.set(`${ref.localName}:${val(ref, 'type')}`, ref);
    }
    effective.set(section.id, new Map(slots));
    parts.set(section.id, part);
    if (section.endingParagraph) {
      const paragraph = part.parentElement?.parentElement;
      if (paragraph?.parentElement !== body)
        throw Error('Moving section breaks inside containers is not supported yet.');
      oldParents.set(section.id, paragraph);
      part.remove();
    }
  }
  if (live.version === 2) {
    for (const inserted of live.inserted) {
      const sourcePart = parts.get(inserted.sourceSectionId);
      if (!sourcePart) throw Error('The authored section has no source properties.');
      parts.set(inserted.id, sourcePart.cloneNode(true) as Element);
      effective.set(inserted.id, new Map(effective.get(inserted.sourceSectionId)));
    }
    for (const override of live.starts) {
      const part = parts.get(override.sectionId);
      if (!part) throw Error('The section start has no source properties.');
      let type = child(part, 'type');
      if (!type) {
        type = part.ownerDocument.createElementNS(WORD_NS, 'w:type');
        const anchor = [...part.children].find(
          (e) =>
            e.namespaceURI === WORD_NS &&
            !['headerReference', 'footerReference', 'footnotePr', 'endnotePr'].includes(
              e.localName,
            ),
        );
        part.insertBefore(type, anchor || null);
      }
      type.setAttributeNS(WORD_NS, 'w:val', override.start);
    }
  }
  return (paragraphs: (Element | undefined)[]) => {
    let changed = materialized || live.breaks.length !== nonFinal.length || live.version === 2;
    const currentSlots = new Map<string, Element>();
    const surviving = [
      ...live.breaks.map((b) => ({ id: b.sectionId, paragraph: b.paragraph })),
      { id: live.finalSectionId, paragraph: null },
    ];
    for (const section of surviving) {
      const part = parts.get(section.id);
      if (!part) continue;
      if (section.paragraph !== null) {
        const target = paragraphs[section.paragraph];
        if (!target || target.parentElement !== body)
          throw Error('The live section break has no supported output paragraph.');
        let pPr = child(target, 'pPr');
        if (!pPr) {
          pPr = target.ownerDocument.createElementNS(WORD_NS, 'w:pPr');
          target.prepend(pPr);
        }
        pPr.insertBefore(part, child(pPr, 'pPrChange') || null);
        changed ||= target !== oldParents.get(section.id);
      }
      const wanted = effective.get(section.id)!;
      const direct = new Map(
        [...part.children]
          .filter((e) => ['headerReference', 'footerReference'].includes(e.localName))
          .map((e) => [`${e.localName}:${val(e, 'type')}`, e]),
      );
      for (const [slot, reference] of wanted) {
        const prior = direct.get(slot) || currentSlots.get(slot);
        const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
        if (prior?.getAttributeNS(R, 'id') !== reference.getAttributeNS(R, 'id')) {
          const anchor = [...part.children].find((e) =>
            reference.localName === 'headerReference'
              ? e.localName !== 'headerReference'
              : !['headerReference', 'footerReference'].includes(e.localName),
          );
          part.insertBefore(reference.cloneNode(true), anchor || null);
          changed = true;
        }
        currentSlots.set(slot, reference);
      }
    }
    return changed;
  };
}
