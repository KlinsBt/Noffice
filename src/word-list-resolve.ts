import type { Node } from '@tiptap/pm/model';
import { isFlatWordList } from './word-flat-list';
import { sectionParagraphs, type WordSectionState } from './word-section-breaks';
import type { WordListDefinition, WordNumbering } from './word-list-layout';

export interface WordListParagraph {
  from: number;
  to: number;
  listFrom: number;
  listTo: number;
  definition: WordListDefinition;
  marker: string;
  value: number;
}
/** Counters belong to source numId, not incidental HTML list wrappers. Empty
 * section-ending paragraphs keep their identity but consume no number. */
export function resolveWordLists(doc: Node, source?: WordNumbering): Map<number, WordListParagraph> | null {
  const family = (value: string) => value.replace(/^(["'])(.*)\1$/, '$2');
  const sources = new Map(source?.paragraphs.map(p => [p.source, p.numbering]));
  const sections = doc.attrs.wordSectionState as WordSectionState | null;
  const paragraphs = sectionParagraphs(doc);
  const hidden = new Set((sections?.breaks || []).map(b => paragraphs[b.paragraph])
    .filter(p => p && !p.node.content.size).map(p => p.from));
  const result = new Map<number, WordListParagraph>(), counters = new Map<string, number>();
  let valid = true;
  doc.forEach((list, listFrom) => {
    if (!['orderedList', 'bulletList'].includes(list.type.name)) return;
    if (!isFlatWordList(list)) { valid = false; return; }
    list.forEach((item, offset) => {
      const p = item.firstChild!, from = listFrom + offset + 2;
      const entry = sources.get(p.attrs.sourceParagraph);
      if (entry?.status !== 'resolved') { valid = false; return; }
      const definition = entry.definition;
      if ((definition.format === 'decimal') !== (list.type.name === 'orderedList') ||
          (list.type.name === 'orderedList' && (list.attrs.type || ![1, definition.start].includes(list.attrs.start))) ||
          p.attrs.indentStart || p.attrs.indentEnd || p.attrs.firstLineIndent ||
          p.attrs.direction === 'rtl' || (p.attrs.textAlign && p.attrs.textAlign !== 'left') ||
          family(p.attrs.paragraphFontFamily || '') !== definition.font || p.attrs.paragraphFontSize !== `${definition.size}pt` ||
          p.attrs.paragraphScript !== 'baseline') { valid = false; return; }
      if (p.attrs.paragraphFontFeatures || p.attrs.paragraphKerning) { valid = false; return; }
      // Marker metrics are currently qualified with matching, plain body fonts.
      // Other typography remains editable in continuous view, never certified.
      p.forEach(run => {
        if (!run.isText || run.marks.some(mark => {
          if (mark.type.name === 'wordEditRun') return false;
          if (mark.type.name !== 'textStyle') return true;
          const a = mark.attrs;
          return (a.fontFamily && family(a.fontFamily) !== definition.font) ||
            (a.fontSize && a.fontSize !== `${definition.size}pt`) || a.color || a.backgroundColor || a.wordFontFeatures || a.wordKerning;
        })) valid = false;
      });
      const value = counters.get(definition.numId) ?? definition.start;
      if (value > 1000000) { valid = false; return; }
      const marker = hidden.has(from) ? '' : definition.format === 'decimal' ? `${value}.` : definition.text;
      result.set(from, { from, to: from + p.nodeSize, listFrom, listTo: listFrom + list.nodeSize, definition, marker, value });
      if (marker) counters.set(definition.numId, value + 1);
    });
  });
  return valid ? result : null;
}
