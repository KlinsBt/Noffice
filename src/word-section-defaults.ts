import { child, wordXml } from './word-xml';
import type { DocxStructure } from './docx-sections';

// Pinned Word16 / de-DE Normal-template defaults. Omitted final properties and
// an explicitly empty final element are different: A4/708twip story distances
// versus Letter/720twip distances. Five independent/control materialization
// pairs preserve all11 native pages. Partial properties and other modes are
// intentionally not filled from this profile. Authored Noffice pages are separate.
const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const margins = 'w:top="1417" w:right="1417" w:bottom="1134" w:left="1417"';
const omitted = `<w:sectPr xmlns:w="${namespace}"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar ${margins} w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>`;
const empty = `<w:sectPr xmlns:w="${namespace}"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar ${margins} w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>`;

export function usesWordSectionDefaults(source?: DocxStructure) {
  const final = source?.sections.at(-1);
  if (source?.compatibility?.mode !== 15 || final?.endingParagraph !== null) return false;
  if (final.propertiesXml === null) return true;
  const root = wordXml(final.propertiesXml).documentElement;
  return root.namespaceURI === namespace && root.localName === 'sectPr' && !root.children.length;
}

/** Return disposable effective properties, leaving retained XML unchanged. */
export function wordFinalSectionProperties(section: Element | null, mode?: number | null) {
  if (mode !== 15 || section?.children.length) return section;
  const resolved = wordXml(section ? empty : omitted).documentElement;
  if (section) for (const attribute of section.attributes)
    resolved.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
  return resolved;
}

/** Materialize only on the exporter's private document. An empty replacement
 * would silently change omitted-section A4 paper to Letter. */
export function materializeWordFinalSection(body: Element, mode?: number | null) {
  const final = child(body, 'sectPr') || null;
  const resolved = wordFinalSectionProperties(final, mode);
  if (!resolved || resolved === final) return false;
  const copy = body.ownerDocument.importNode(resolved, true);
  if (final) final.replaceWith(copy);
  else body.append(copy);
  return true;
}
