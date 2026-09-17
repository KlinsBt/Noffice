import { Extension, type Command } from '@tiptap/core';
import type { EditorState } from '@tiptap/pm/state';
import type { Mark } from '@tiptap/pm/model';
import { closeHistory } from '@tiptap/pm/history';

export type WordScript = 'baseline' | 'superscript' | 'subscript';
export function wordScriptValue(value: unknown): WordScript | null {
  return value === 'baseline' || value === 'superscript' || value === 'subscript' ? value : null;
}
const scriptOf = (marks: readonly Mark[]): WordScript => marks.some(m => m.type.name === 'superscript')
  ? 'superscript' : marks.some(m => m.type.name === 'subscript') ? 'subscript' : 'baseline';

/** Inline selection supplies the aggregate; mark-only selections use persistent
 * paragraph formatting. A stored typing override precedes an empty mark. */
export function wordSelectionScript(state: EditorState): WordScript | null {
  const { from, to, empty, $from } = state.selection;
  if (empty) return state.storedMarks ? scriptOf(state.storedMarks)
    : !$from.parent.content.size ? wordScriptValue($from.parent.attrs.paragraphScript) || 'baseline'
    : scriptOf($from.marks());
  const inline: WordScript[] = [], marks: WordScript[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.isInline) inline.push(scriptOf(node.marks));
    else if (node.isTextblock && to > pos + 1 + node.content.size)
      marks.push(wordScriptValue(node.attrs.paragraphScript) || 'baseline');
  });
  const values = inline.length ? inline : marks;
  return !values.length ? 'baseline' : values.every(value => value === values[0]) ? values[0] : null;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    wordScript: {
      setWordScript: (value: WordScript) => ReturnType;
      toggleWordScript: (value: 'superscript' | 'subscript') => ReturnType;
    };
  }
}

function setScript(value: WordScript): Command {
  return ({ tr, dispatch }) => {
    if (!wordScriptValue(value)) return false;
    if (!dispatch) return true;
    const { from, to, empty, $from } = tr.selection;
    const superType = tr.doc.type.schema.marks.superscript, subType = tr.doc.type.schema.marks.subscript;
    const wanted = value === 'baseline' ? null : value === 'superscript' ? superType : subType;
    let left = from, right = to;
    if (empty && $from.parent.isTextblock) {
      const text = $from.parent.textBetween(0, $from.parent.content.size, '\ufffc', '\ufffc');
      const offset = $from.parentOffset;
      // Same bounded ASCII interior-word semantics as explicit font commands.
      if (/[a-z]/i.test(text[offset - 1] || '') && /[a-z]/i.test(text[offset] || '')) {
        let a = offset, b = offset;
        while (a > 0 && /[a-z]/i.test(text[a - 1])) a--;
        while (b < text.length && /[a-z]/i.test(text[b])) b++;
        left = $from.start() + a; right = $from.start() + b;
      }
    }
    closeHistory(tr);
    tr.doc.nodesBetween(left, right, (node, pos) => {
      if (!node.isInline) return;
      const a = Math.max(left, pos), b = Math.min(right, pos + node.nodeSize);
      if (a >= b) return;
      if (superType.isInSet(node.marks) && wanted !== superType) tr.removeMark(a, b, superType);
      if (subType.isInSet(node.marks) && wanted !== subType) tr.removeMark(a, b, subType);
      if (wanted && !wanted.isInSet(node.marks)) tr.addMark(a, b, wanted.create());
    });
    tr.doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isTextblock) return;
      const start = pos + 1, end = start + node.content.size;
      const selected = empty ? from === end : to > end || (node.content.size > 0 && from <= start && to >= end);
      if (selected && (wordScriptValue(node.attrs.paragraphScript) || 'baseline') !== value)
        tr.setNodeAttribute(pos, 'paragraphScript', value);
      return false;
    });
    if (empty) {
      const marks = (tr.storedMarks || $from.marks()).filter(mark => mark.type !== superType && mark.type !== subType);
      tr.setStoredMarks(wanted ? [...marks, wanted.create()] : marks);
    }
    return true;
  };
}

export const WordScriptFormatting = Extension.create({
  name: 'wordScriptFormatting',
  addGlobalAttributes() {
    return [{ types: ['paragraph', 'heading'], attributes: {
      paragraphScript: {
        default: null,
        parseHTML: element => wordScriptValue(element.getAttribute('data-word-paragraph-script')),
        renderHTML: attrs => {
          const value = wordScriptValue(attrs.paragraphScript);
          return value ? { 'data-word-paragraph-script': value } : {};
        },
      },
    } }];
  },
  addCommands() {
    return {
      setWordScript: setScript,
      toggleWordScript: value => ({ state, commands }) =>
        commands.setWordScript(wordSelectionScript(state) === value ? 'baseline' : value),
    };
  },
});
