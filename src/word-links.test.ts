import { expect, it } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { selectedWordLink, applyWordLink } from './word-links';

it.each([false, true])(
  'captures an immediate native text selection before opening a link dialog (backward: %s)',
  (backward) => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const editor = new Editor({
      element: host,
      editorProps: { handleScrollToSelection: () => true },
      extensions: wordExtensions(),
      content: '<p><strong>Alpha</strong><em> beta</em></p>',
    });
    try {
      editor.view.dom.tabIndex = 0;
      editor.view.focus();
      const first = editor.view.dom.querySelector('strong')!.firstChild!,
        last = editor.view.dom.querySelector('em')!.firstChild!;
      document
        .getSelection()!
        .setBaseAndExtent(
          backward ? last : first,
          backward ? 5 : 0,
          backward ? first : last,
          backward ? 0 : 5,
        );
      expect(editor.state.selection.empty).toBe(true);
      const selected = selectedWordLink(editor);
      expect(selected.text).toBe('Alpha beta');
      expect(editor.state.selection.anchor).toBe(backward ? 11 : 1);
      applyWordLink(editor, 'https://example.com/selection', selected.text, selected.text);
      expect(editor.view.dom.querySelector('a')?.textContent).toBe('Alpha beta');
      expect(editor.getHTML()).toContain('<strong>');
      expect(editor.getHTML()).toContain('<em>');
    } finally {
      editor.destroy();
      host.remove();
      document.getSelection()?.removeAllRanges();
    }
  },
);

it('adds a link across mixed formatting, edits from the cursor and removes it with independent undo steps', () => {
  const editor = new Editor({
    extensions: wordExtensions(),
    content: '<p><strong>Alpha</strong><em> beta</em></p>',
  });
  try {
    editor.commands.setTextSelection({ from: 1, to: 11 });
    const initial = selectedWordLink(editor);
    expect(initial.text).toBe('Alpha beta');
    applyWordLink(editor, 'https://example.com', initial.text, initial.text);
    expect(editor.getHTML()).toContain('<strong>');
    expect(editor.getHTML()).toContain('<em>');
    editor.commands.setTextSelection(3);
    expect(selectedWordLink(editor).text).toBe('Alpha beta');
    applyWordLink(editor, 'mailto:team@example.com', initial.text, initial.text);
    expect(editor.getHTML()).toContain('mailto:team@example.com');
    editor.commands.undo();
    expect(editor.getHTML()).toContain('https://example.com');
    editor.commands.redo();
    selectedWordLink(editor);
    applyWordLink(editor, null, initial.text, initial.text);
    expect(editor.getHTML()).not.toContain('<a ');
    expect(editor.getHTML()).toContain('<strong>');
    editor.commands.undo();
    expect(editor.getHTML()).toContain('mailto:team@example.com');
  } finally {
    editor.destroy();
  }
});
it('inserts display text at an empty cursor without leaking a link into subsequent typing or accepting unsafe schemes', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p>Start&nbsp;</p>' });
  try {
    editor.commands.setTextSelection(7);
    const before = editor.getHTML();
    expect(() => applyWordLink(editor, 'javascript:alert(1)', 'Bad', '')).toThrow();
    expect(editor.getHTML()).toBe(before);
    applyWordLink(editor, 'https://example.com/', 'Visit', '');
    editor.commands.insertContent(' after');
    const dom = new DOMParser().parseFromString(editor.getHTML(), 'text/html');
    expect(dom.querySelector('a')?.textContent).toBe('Visit');
    expect(dom.body.textContent).toBe('Start\u00a0Visit after');
    editor.commands.setTextSelection(9);
    const selected = selectedWordLink(editor);
    applyWordLink(editor, 'https://example.com/next', 'Renamed', selected.text);
    expect(editor.getText()).toBe('Start\u00a0Renamed after');
  } finally {
    editor.destroy();
  }
});
