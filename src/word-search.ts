import type { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { textMatches, type SearchOptions, type TextMatch } from './text-search';

export function wordMatches(
  editor: Editor,
  query: string,
  options: SearchOptions = {},
): TextMatch[] {
  const found: TextMatch[] = [],
    budget = { characters: 0, matches: 0 };
  editor.state.doc.descendants((node, pos) => {
    if (!node.isTextblock) return;
    let text = '',
      start = pos + 1;
    const flush = () => {
      found.push(
        ...textMatches(text, query, options, budget).map((m) => ({
          from: start + m.from,
          to: start + m.to,
        })),
      );
      text = '';
    };
    node.forEach((child, offset) => {
      if (child.isText) {
        if (!text) start = pos + 1 + offset;
        text += child.text;
      } else flush();
    });
    flush();
    return false;
  });
  return found;
}

export function replaceWordMatches(editor: Editor, matches: TextMatch[], replacement: string) {
  if (!matches.length) return;
  if (
    editor.state.doc.textContent.length +
      matches.reduce((n, m) => n + replacement.length - (m.to - m.from), 0) >
    2000000
  )
    throw Error('Replacement would exceed the document text limit.');
  const tr = closeHistory(editor.state.tr);
  for (const match of [...matches].reverse()) {
    const marks = editor.state.doc.nodeAt(match.from)?.marks || [];
    if (replacement) tr.replaceWith(match.from, match.to, editor.schema.text(replacement, marks));
    else tr.delete(match.from, match.to);
  }
  editor.view.dispatch(tr.scrollIntoView());
  editor.view.dispatch(closeHistory(editor.state.tr));
}
