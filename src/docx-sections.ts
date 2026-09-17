import { z } from 'zod';
import { wordCompatibilitySchema } from './word-compatibility';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const children = (e: Element, name: string) =>
  [...e.children].filter((n) => n.namespaceURI === W && n.localName === name);
const attr = (e: Element, name: string) => e.getAttributeNS(W, name);
const partRef = z.object({ type: z.string().max(100), relationshipId: z.string().max(200) });
export const docxStructureSchema = z.object({
  version: z.literal(1),
  sourcePath: z.literal('word/document.xml'),
  compatibility: wordCompatibilitySchema.optional(),
  sections: z
    .array(
      z.object({
        id: z.string().regex(/^word\/document\.xml#section:\d+$/),
        endingParagraph: z
          .string()
          .regex(/^0:\d+$/)
          .nullable(),
        paragraphs: z.array(z.string().regex(/^0:\d+$/)).max(100000),
        // Original section properties, not the effective layout after subsequent commands.
        propertiesXml: z.string().max(100000).nullable(),
        headers: z.array(partRef).max(10),
        footers: z.array(partRef).max(10),
        drawings: z
          .array(
            z.object({
              paragraph: z.string(),
              id: z.string().max(200),
              relationshipIds: z.array(z.string().max(200)).max(100),
            }),
          )
          .max(10000),
        notes: z
          .array(
            z.object({
              paragraph: z.string(),
              kind: z.enum(['footnote', 'endnote']),
              id: z.string().max(100),
            }),
          )
          .max(100000),
      }),
    )
    .min(1)
    .max(10000),
});
export type DocxStructure = z.infer<typeof docxStructureSchema>;

/** Current body sections only. Revision history and textbox/table-local markup are not sections. */
export function liveSectionProperties(body: Element): Element[] {
  return [...body.getElementsByTagNameNS(W, 'sectPr')].filter((section) => {
    if (section.parentElement === body) return true;
    if (
      section.parentElement?.localName !== 'pPr' ||
      section.parentElement.parentElement?.localName !== 'p'
    )
      return false;
    for (
      let p = section.parentElement.parentElement.parentElement;
      p && p !== body;
      p = p.parentElement
    )
      if (
        p.namespaceURI === W &&
        ['tbl', 'txbxContent', 'sectPrChange', 'pPrChange', 'del', 'moveFrom'].includes(p.localName)
      )
        return false;
    return true;
  });
}

export function readDocxStructure(main: Document): DocxStructure {
  const body = [...main.getElementsByTagNameNS(W, 'body')][0];
  if (!body) throw Error('The Word document has no body.');
  const paragraphs = [...main.getElementsByTagNameNS(W, 'p')];
  const index = new Map(paragraphs.map((p, i) => [p, `0:${i}`]));
  const positions = new Map(paragraphs.map((p, i) => [p, i]));
  const properties = liveSectionProperties(body);
  const final = properties.find((p) => p.parentElement === body);
  const breaks = properties.filter((p) => p !== final);
  let start = 0;
  const sections: DocxStructure['sections'] = [];
  for (const section of [...breaks, final]) {
    const ending = section && section !== final ? section.parentElement!.parentElement! : null;
    const nested = ending ? [...ending.getElementsByTagNameNS(W, 'p')] : [];
    const end = ending ? positions.get(nested.at(-1) || ending)! : paragraphs.length - 1;
    const members = paragraphs.slice(start, end + 1);
    const references = (name: string) =>
      section
        ? children(section, name).map((e) => ({
            type: attr(e, 'type') || 'default',
            relationshipId: e.getAttributeNS(R, 'id') || '',
          }))
        : [];
    sections.push({
      id: `word/document.xml#section:${sections.length}`,
      endingParagraph: ending ? index.get(ending)! : null,
      paragraphs: members.map((p) => index.get(p)!),
      propertiesXml: section ? new XMLSerializer().serializeToString(section) : null,
      headers: references('headerReference'),
      footers: references('footerReference'),
      drawings: members.flatMap((p) =>
        [...p.getElementsByTagNameNS(WP, 'anchor')]
          .filter((a) => closestParagraph(a) === p)
          .map((anchor) => ({
            paragraph: index.get(p)!,
            id: anchor.getElementsByTagNameNS(WP, 'docPr')[0]?.getAttribute('id') || '',
            relationshipIds: [...anchor.getElementsByTagNameNS(A, 'blip')].map(
              (b) => b.getAttributeNS(R, 'embed') || b.getAttributeNS(R, 'link') || '',
            ),
          })),
      ),
      notes: members.flatMap((p) =>
        (['footnote', 'endnote'] as const).flatMap((kind) =>
          [...p.getElementsByTagNameNS(W, kind + 'Reference')]
            .filter((n) => closestParagraph(n) === p)
            .map((n) => ({ paragraph: index.get(p)!, kind, id: attr(n, 'id') || '' })),
        ),
      ),
    });
    start = end + 1;
  }
  return docxStructureSchema.parse({ version: 1, sourcePath: 'word/document.xml', sections });
}
function closestParagraph(node: Element): Element | null {
  for (let p = node.parentElement; p; p = p.parentElement)
    if (p.namespaceURI === W && p.localName === 'p') return p;
  return null;
}
