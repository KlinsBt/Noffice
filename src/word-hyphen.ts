import { Node } from '@tiptap/core';
import { Plugin, PluginKey, type Transaction } from '@tiptap/pm/state';
import { Fragment, type Node as EditorNode } from '@tiptap/pm/model';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    wordHyphen: { insertWordHyphen: (kind: 'optional' | 'literal') => ReturnType };
  }
}

/** Word's optional control and literal U+00AD are different characters. The
 * native value is used only in retained-XML comparisons, never DOM paint. */
export function wordHyphenText(kind: unknown, mode: 'display' | 'native' | 'plain' = 'display') {
  if (kind === 'optional') return mode === 'native' ? '\u001f' : '\u00ad';
  if (kind === 'literal') return mode === 'display' ? '-' : '\u00ad';
  return null;
}

/** Normalize only a committed text insertion. Word stores a literal U+00AD
 * separately from its optional control; preserve the insertion's resolved
 * marks and positions in the same transaction. Existing text is untouched. */
export function literalWordInput(tr: Transaction, from: number, to: number) {
  const type = tr.doc.type.schema.nodes.wordHyphen;
  if (!type) return tr;
  const inserted = tr.doc.slice(from, to), nodes: EditorNode[] = [];
  if (inserted.openStart || inserted.openEnd) return tr;
  let changed = false;
  inserted.content.forEach(node => {
    if (!node.isText || !node.text?.includes('\u00ad')) { nodes.push(node); return; }
    changed = true;
    const parts = node.text.split('\u00ad');
    parts.forEach((text, index) => {
      if (index) nodes.push(type.create({ kind: 'literal' }, null, node.marks));
      if (text) nodes.push(tr.doc.type.schema.text(text, node.marks));
    });
  });
  // One replacement keeps transaction mapping linear for multi-character input.
  if (changed) tr.replaceWith(from, to, Fragment.fromArray(nodes));
  return tr;
}

export const WordHyphen = Node.create({
  name: 'wordHyphen',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: false,
  addAttributes: () => ({ kind: { default: 'optional', rendered: false } }),
  parseHTML: () => [{
    tag: 'span[data-word-hyphen]',
    getAttrs: element => {
      const kind = element.getAttribute('data-word-hyphen'), text = wordHyphenText(kind);
      return text !== null && !element.children.length && element.textContent === text ? { kind } : false;
    },
  }],
  renderHTML: ({ node }) => {
    const text = wordHyphenText(node.attrs.kind);
    if (text === null) throw Error('Invalid Word hyphen character.');
    return ['span', { 'data-word-hyphen': node.attrs.kind }, text];
  },
  renderText: ({ node }) => wordHyphenText(node.attrs.kind, 'plain') || '',
  addProseMirrorPlugins() {
    return [new Plugin({
      key: new PluginKey('wordLiteralHyphenInput'),
      props: {
        handleTextInput(view, from, to, text) {
          if (!view.editable || view.composing || !text.includes('\u00ad') || /[\r\n\t]/.test(text)) return false;
          // The existing paragraph-input handler owns selections that include
          // a paragraph mark; it applies the same character normalization.
          if (!view.state.doc.resolve(from).sameParent(view.state.doc.resolve(to))) return false;
          view.dispatch(literalWordInput(view.state.tr.insertText(text, from, to), from, from + text.length).scrollIntoView());
          return true;
        },
      },
    })];
  },
  addCommands() {
    return {
      insertWordHyphen: kind => ({ tr, commands }) => {
        if (wordHyphenText(kind) === null) return false;
        return commands.insertContent({ type: this.name, attrs: { kind },
          marks: (tr.storedMarks || tr.selection.$from.marks()).map(mark => mark.toJSON()) });
      },
    };
  },
  addKeyboardShortcuts() {
    // Pinned Windows Word optional-hyphen shortcut. Native physical-keyboard
    // acceptance remains distinct from browser key handling and DOCX semantics.
    return { 'Ctrl--': () => this.editor.commands.insertWordHyphen('optional') };
  },
});
