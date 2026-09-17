import { Extension } from '@tiptap/core';
import { wordKerningValue } from './word-kerning';
import { wordFontFeaturesValue } from './word-font-features';
import { wordScriptValue } from './word-script';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { Mapping } from '@tiptap/pm/transform';
import { isHistoryTransaction } from '@tiptap/pm/history';

/** Empty paragraph marks carry the font for newly entered text. Materialize that
 * font on the text in the same history transaction so retained DOCX agrees. */
export const WordFontInput = Extension.create({
  name: 'wordFontInput',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('wordFontInput'),
        appendTransaction(transactions, previous, next) {
          const cursor = previous.selection.$from;
          const terminalLine =
            cursor.parentOffset === cursor.parent.content.size &&
            ['hardBreak', 'wordPageBreak', 'wordColumnBreak'].includes(
              cursor.parent.lastChild?.type.name || '',
            );
          if (
            !transactions.some((t) => t.docChanged) ||
            transactions.some(isHistoryTransaction) ||
            !previous.selection.empty ||
            cursor.depth < 1 ||
            (cursor.parent.content.size && !terminalLine) ||
            (terminalLine && transactions.some((t) => t.getMeta('uiEvent') === 'paste')) ||
            !['paragraph', 'heading'].includes(cursor.parent.type.name)
          )
            return null;
          const mapping = new Mapping();
          for (const tr of transactions) mapping.appendMapping(tr.mapping);
          const start = mapping.map(cursor.before(), -1);
          const paragraph = next.doc.nodeAt(start);
          const markType = next.schema.marks.textStyle;
          if (!paragraph?.isTextblock || !markType) return null;
          const { paragraphFontSize, paragraphFontFamily } = paragraph.attrs;
          const paragraphKerning = wordKerningValue(paragraph.attrs.paragraphKerning);
          const features = wordFontFeaturesValue(paragraph.attrs.paragraphFontFeatures);
          const script = !terminalLine && previous.storedMarks === null &&
            !transactions.some(t => t.getMeta('uiEvent') === 'paste')
            ? wordScriptValue(paragraph.attrs.paragraphScript) : null;
          if (!paragraphFontSize && !paragraphFontFamily && paragraphKerning === null && features === null && !script) return null;
          const tr = next.tr;
          const insertedFrom = mapping.map(previous.selection.from, -1);
          const insertedTo = mapping.map(previous.selection.to, 1);
          paragraph.forEach((node, offset) => {
            if (!node.isText && node.type.name !== 'wordHyphen') return;
            const from = terminalLine
              ? Math.max(start + 1 + offset, insertedFrom)
              : start + 1 + offset;
            const to = terminalLine
              ? Math.min(start + 1 + offset + node.nodeSize, insertedTo)
              : start + 1 + offset + node.nodeSize;
            if (from >= to) return;
            if (script && script !== 'baseline' &&
                !node.marks.some(mark => ['superscript', 'subscript'].includes(mark.type.name))) {
              const type = next.schema.marks[script];
              if (type) tr.addMark(from, to, type.create());
            }
            const old = markType.isInSet(node.marks);
            const attrs = { ...old?.attrs };
            if (terminalLine && !markType.isInSet(previous.storedMarks || [])) {
              const inherited = markType.isInSet(cursor.marks());
              if (attrs.fontSize === inherited?.attrs.fontSize && paragraphFontSize)
                attrs.fontSize = paragraphFontSize;
              if (attrs.fontFamily === inherited?.attrs.fontFamily && paragraphFontFamily)
                attrs.fontFamily = paragraphFontFamily;
              if (attrs.wordKerning === inherited?.attrs.wordKerning && paragraphKerning !== null)
                attrs.wordKerning = paragraphKerning;
              if (attrs.wordFontFeatures === inherited?.attrs.wordFontFeatures && features !== null)
                attrs.wordFontFeatures = features;
            }
            if (!attrs.fontSize && paragraphFontSize) attrs.fontSize = paragraphFontSize;
            if (!attrs.fontFamily && paragraphFontFamily) attrs.fontFamily = paragraphFontFamily;
            if (attrs.wordKerning == null && paragraphKerning !== null)
              attrs.wordKerning = paragraphKerning;
            if (attrs.wordFontFeatures == null && features !== null) attrs.wordFontFeatures = features;
            const mark = markType.create(attrs);
            if (!old?.eq(mark)) tr.addMark(from, to, mark);
          });
          return tr.steps.length ? tr : null;
        },
      }),
    ];
  },
});
