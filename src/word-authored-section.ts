import type { Node } from '@tiptap/pm/model';
import type { WordContent } from './model';
import type { WordSectionLayout } from './word-section-layout';
import type { WordSectionMap } from './word-section-ranges';

/** A document authored without retained OOXML has one section with explicit
 * global page settings. This view identity is never a retained source identity. */
export function authoredWordSection(content: WordContent): WordSectionLayout | null {
  if (content.docxStructure || content.sectionState || /data-source-paragraph/i.test(content.html))
    return null;
  const [short, long] = content.paper === 'letter' ? [12240, 15840] : [11906, 16838];
  const orientation = content.orientation || 'portrait';
  const margin = content.margin === 'narrow' ? 720 : content.margin === 'wide' ? 2160 : 1440;
  const options = content.stories?.sectionOptions?.find((o) => o.sectionId === 'authored-body');
  const section: WordSectionLayout = {
    id: 'authored-body',
    paragraphs: [],
    width: orientation === 'landscape' ? long : short,
    height: orientation === 'landscape' ? short : long,
    orientation,
    margins: {
      left: margin,
      right: margin,
      top: margin,
      bottom: margin,
      header: options?.headerDistance ?? 708,
      footer: options?.footerDistance ?? 708,
      gutter: 0,
    },
    headers: { default: null, even: null, first: null },
    footers: { default: null, even: null, first: null },
    differentFirstPage: options?.differentFirstPage ?? false,
    simpleBody: true,
    start: 'nextPage',
  };
  for (const ref of content.stories?.references || []) {
    if (ref.sectionId !== section.id || ref.linked) continue;
    section[ref.kind === 'header' ? 'headers' : 'footers'][ref.slot] = {
      relationshipId: ref.relationshipId,
      sourceSectionId: section.id,
      inherited: false,
    };
  }
  return section;
}

/** Map live positions independently of imported source-boundary repair. */
export function authoredWordSectionMap(
  doc: Node,
  section: WordSectionLayout,
): WordSectionMap | undefined {
  let valid = !doc.attrs.wordSectionState;
  const paragraphs: WordSectionMap['paragraphs'] = [];
  doc.descendants((node, from) => {
    if (node.attrs.sourceParagraph) valid = false;
    if (!['paragraph', 'heading'].includes(node.type.name)) return;
    paragraphs.push({ from, to: from + node.nodeSize, sourceId: null, sectionId: section.id });
    return false;
  });
  return valid
    ? {
        size: doc.content.size,
        ranges: [{ id: section.id, from: 0, to: doc.content.size }],
        paragraphs,
        issues: [],
      }
    : undefined;
}
