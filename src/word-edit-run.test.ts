import { Editor, type JSONContent } from '@tiptap/core';
import { expect, it } from 'vitest';
import { closeHistory } from '@tiptap/pm/history';
import { wordExtensions, wordJSON } from './word-extensions';
import { sanitizeHTML } from './formats';
import { checkpointWordEditSession, initializeWordEditSession } from './word-edit-run';

it('keeps the first saved session atomic with editing and restores it across reload', () => {
  let editor = new Editor({ extensions: wordExtensions({ retainedEditRuns: true }), content: '<p data-source-paragraph="p0">ABC</p>' });
  try {
    initializeWordEditSession(editor);
    checkpointWordEditSession(editor.view);
    expect(editor.state.doc.attrs.wordInitialEditSession).toBeNull();
    editor.commands.insertContentAt({ from: 2, to: 3 }, 'X');
    const initial = editor.state.doc.attrs.wordInitialEditSession;
    const first = editor.state.doc.firstChild!.child(1).marks.find(mark => mark.type.name === 'wordEditRun')!.attrs.session;
    expect(initial).toBe(first);
    const saved = editor.getHTML();
    editor.commands.undo();
    expect(editor.getText()).toBe('ABC');
    expect(editor.state.doc.attrs.wordInitialEditSession).toBeNull();
    editor.commands.redo();
    expect(editor.state.doc.attrs.wordInitialEditSession).toBe(initial);
    checkpointWordEditSession(editor.view);
    expect(editor.state.doc.attrs.wordInitialEditSession).toBe(initial);
    editor.destroy();
    editor = new Editor({ extensions: wordExtensions({ retainedEditRuns: true }), content: saved });
    initializeWordEditSession(editor, initial);
    expect(editor.can().undo()).toBe(false);
    editor.commands.insertContentAt(4, 'Y');
    expect(editor.state.doc.attrs.wordInitialEditSession).toBe(initial);
    const last = editor.state.doc.firstChild!.lastChild!.marks.find(mark => mark.type.name === 'wordEditRun')!.attrs.session;
    expect(last).not.toBe(initial);
    editor.commands.undo();
    expect(editor.getHTML()).toBe(saved);
    expect(editor.state.doc.attrs.wordInitialEditSession).toBe(initial);
    expect(saved).not.toContain('wordInitialEditSession');
  } finally { editor.destroy(); }
});

it('starts a new save session without changing the document or adding an undo event', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p></p>' });
  try {
    editor.commands.insertContent('before');
    const saved = editor.getJSON(),
      selection = editor.state.selection.toJSON();
    checkpointWordEditSession(editor.view);
    expect(editor.getJSON()).toEqual(saved);
    expect(editor.state.selection.toJSON()).toEqual(selection);
    editor.commands.insertContent(' added');
    const pieces = (editor.getJSON() as JSONContent).content![0].content!;
    expect(pieces.map((p) => p.text)).toEqual(['before', ' added']);
    expect(pieces[0].marks).not.toEqual(pieces[1].marks);
    editor.commands.undo();
    expect(editor.getJSON()).toEqual(saved);
    editor.commands.undo();
    expect(editor.getText()).toBe('');
    expect(editor.can().undo()).toBe(false);
    editor.commands.redo();
    editor.commands.redo();
    expect(editor.getJSON().content![0].content).toEqual(pieces);
  } finally {
    editor.destroy();
  }
});

it('keeps new-paragraph edit sessions through history, sanitized storage and reload', () => {
  let editor = new Editor({ extensions: wordExtensions(), content: '<p>er</p>' });
  try {
    editor.commands.setTextSelection(3);
    editor.commands.insertContent('!');
    editor.commands.insertContent('?');
    const content = (editor.getJSON() as JSONContent).content![0].content!;
    expect(content.map((n) => n.text)).toEqual(['er', '!?']);
    const session = content[1].marks!.find((m) => m.type === 'wordEditRun')!.attrs!.session;
    expect(session).toMatch(/^[a-f\d]{32}$/);
    editor.commands.undo();
    expect(editor.getText()).toBe('er');
    editor.commands.redo();
    expect(editor.getJSON().content![0].content).toEqual(content);
    const saved = sanitizeHTML(editor.getHTML());
    expect(wordJSON(saved).content![0].content).toEqual(content);
    editor.destroy();
    editor = new Editor({ extensions: wordExtensions(), content: saved });
    editor.commands.setTextSelection(5);
    editor.commands.insertContent('+');
    const reloaded = (editor.getJSON() as JSONContent).content![0].content!;
    expect(reloaded.map((n) => n.text)).toEqual(['er', '!?', '+']);
    expect(reloaded[2].marks!.find((m) => m.type === 'wordEditRun')!.attrs!.session).not.toBe(
      session,
    );
  } finally {
    editor.destroy();
  }
});

it('preserves retained paragraph mapping and excludes structural splits/deletion/history', () => {
  const editor = new Editor({
    extensions: wordExtensions(),
    content: '<p data-source-paragraph="p0">before</p>',
  });
  try {
    editor.commands.setTextSelection(7);
    editor.commands.insertContent('!');
    expect(editor.getHTML()).not.toContain('data-word-edit-run');
    editor.view.dispatch(closeHistory(editor.state.tr));
    editor.commands.setTextSelection(4);
    editor.commands.splitBlock();
    expect(editor.getHTML()).not.toContain('data-word-edit-run');
    editor.commands.undo();
    expect(editor.getText()).toBe('before!');
    editor.commands.setTextSelection({ from: 1, to: 4 });
    editor.commands.deleteSelection();
    expect(editor.getHTML()).not.toContain('data-word-edit-run');
  } finally {
    editor.destroy();
  }
});

it('does not accept unbounded or executable session metadata', () => {
  for (const session of ['javascript:alert(1)', 'f'.repeat(33), '', 'xyz']) {
    expect(
      wordJSON(`<p><span data-word-edit-run="${session}">safe</span></p>`).content![0].content![0]
        .marks,
    ).toBeUndefined();
  }
});

it('distinguishes a retained story join from replacing the same cross-paragraph text', () => {
  const editor = new Editor({
    extensions: wordExtensions({ retainedEditRuns: true }),
    content: '<p data-source-paragraph="p0">first</p><p data-source-paragraph="p1">second</p>',
  });
  try {
    editor.commands.setTextSelection(editor.state.doc.firstChild!.nodeSize + 1);
    expect(editor.commands.joinBackward()).toBe(true);
    expect(editor.getText()).toBe('firstsecond');
    expect(editor.getHTML()).not.toContain('data-word-edit-run');
    editor.commands.undo();
    editor.commands.setTextSelection({ from: 1, to: editor.state.doc.content.size - 1 });
    editor.commands.insertContent('firstsecond');
    expect(editor.getText()).toBe('firstsecond');
    expect(editor.getHTML()).toContain('data-word-edit-run');
    const saved = editor.getHTML();
    editor.commands.undo();
    expect(editor.state.doc.childCount).toBe(2);
    expect(editor.getHTML()).not.toContain('data-word-edit-run');
    editor.commands.redo();
    expect(editor.getHTML()).toBe(saved);
  } finally {
    editor.destroy();
  }
});

it('retains composition provenance through DOM wrapper removal, candidate updates and history', () => {
  const editor = new Editor({ extensions: wordExtensions({ retainedEditRuns: true }),
    content: '<p data-source-paragraph="p0">Original text</p>' });
  try {
    const original = editor.getJSON();
    editor.view.dispatch(editor.state.tr.insertText('Replace', 1, 9).setMeta('composition', 1));
    const marker = editor.state.doc.firstChild!.firstChild!.marks.find(mark => mark.type.name === 'wordEditRun')!;
    expect(marker).toBeDefined();
    editor.view.dispatch(editor.state.tr.removeMark(1, 8, marker).setMeta('composition', 1));
    editor.view.dispatch(editor.state.tr.insertText('ment', 8, 8).setMeta('composition', 1));
    editor.view.dispatch(editor.state.tr.removeMark(8, 12, marker).setMeta('composition', 1));
    const paragraph = editor.state.doc.firstChild!;
    expect(paragraph.childCount).toBe(2);
    expect(paragraph.firstChild!.text).toBe('Replacement');
    expect(paragraph.firstChild!.marks).toContainEqual(marker);
    expect(paragraph.lastChild!.text).toBe(' text');
    expect(paragraph.lastChild!.marks).not.toContainEqual(marker);
    const final = editor.getJSON();
    editor.commands.undo(); expect(editor.getJSON()).toEqual(original);
    editor.commands.redo(); expect(editor.getJSON()).toEqual(final);
    // Explicit formatting commands remain distinct from composition parsing.
    editor.view.dispatch(editor.state.tr.removeMark(1, 12, marker));
    expect(editor.getHTML()).not.toContain('data-word-edit-run');
  } finally { editor.destroy(); }
});

it('cancels composed text without transferring its save provenance to retained text', () => {
  const editor = new Editor({ extensions: wordExtensions({ retainedEditRuns: true }), content: '<p data-source-paragraph="p0">Original</p>' });
  try {
    const original = editor.getJSON();
    editor.view.dispatch(editor.state.tr.insertText('ni', 1).setMeta('composition', 1));
    const marker = editor.state.doc.firstChild!.firstChild!.marks.find(mark => mark.type.name === 'wordEditRun')!;
    editor.view.dispatch(editor.state.tr.removeMark(1, 3, marker).delete(1, 3).setMeta('composition', 1));
    expect(editor.getJSON()).toEqual(original);
  } finally { editor.destroy(); }
});
