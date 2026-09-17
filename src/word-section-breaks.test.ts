import { it, expect } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import {
  WordEditorSections,
  updateWordSectionSource,
  wordSectionMap,
} from './word-editor-sections';
import {
  initializeSectionState,
  validateSectionState,
  wordSectionStateSchema,
} from './word-section-breaks';
import type { DocxStructure } from './docx-sections';

const source: DocxStructure = {
  version: 1,
  sourcePath: 'word/document.xml',
  sections: [
    {
      id: 'word/document.xml#section:0',
      endingParagraph: '0:0',
      paragraphs: ['0:0'],
      propertiesXml: null,
      headers: [],
      footers: [],
      notes: [],
      drawings: [],
    },
    {
      id: 'word/document.xml#section:1',
      endingParagraph: null,
      paragraphs: ['0:1'],
      propertiesXml: null,
      headers: [],
      footers: [],
      notes: [],
      drawings: [],
    },
  ],
};
const make = (first = '<p data-source-paragraph="0:0">AlphaBeta</p>') => {
  const editor = new Editor({
    extensions: [...wordExtensions(), WordEditorSections.configure({ source })],
    content: first + '<p data-source-paragraph="0:1">Gamma</p>',
  });
  updateWordSectionSource(editor, source);
  return editor;
};
it.each(['', 'AlphaBeta'])(
  'retains the section boundary when formatting a paragraph containing "%s"',
  (text) => {
    const editor = make(`<p data-source-paragraph="0:0">${text}</p>`);
    try {
      const before = editor.getJSON();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(0, undefined, {
          ...editor.state.doc.firstChild!.attrs,
          textAlign: 'right',
        }),
      );
      expect(editor.state.doc.attrs.wordSectionState).toEqual(before.attrs!.wordSectionState);
      editor.commands.undo();
      expect(editor.getJSON()).toEqual(before);
    } finally {
      editor.destroy();
    }
  },
);
it.each(['page', 'column'] as const)(
  'maps a section-ending paragraph through native %s insertion and Undo',
  (kind) => {
    const editor = make();
    try {
      const original = editor.getJSON();
      editor.commands.setTextSelection(6);
      editor.commands.insertWordFlowBreak(kind);
      const count = kind === 'page' ? 3 : 2;
      expect(editor.state.doc.attrs.wordSectionState.breaks[0].paragraph).toBe(count - 1);
      expect(wordSectionMap(editor.state)?.paragraphs.map((p) => p.sectionId)).toEqual([
        ...Array(count).fill(source.sections[0].id),
        source.sections[1].id,
      ]);
      const inserted = editor.getJSON();
      editor.commands.undo();
      expect(editor.getJSON()).toEqual(original);
      editor.commands.redo();
      expect(editor.getJSON()).toEqual(inserted);
    } finally {
      editor.destroy();
    }
  },
);
it('retains the section end after a leading page break while moving the source anchor before it', () => {
  const editor = make();
  try {
    const original = editor.getJSON();
    editor.commands.setTextSelection(1);
    editor.commands.insertWordFlowBreak('page');
    expect(editor.state.doc.child(0).attrs.sourceParagraph).toBe('0:0');
    expect(editor.state.doc.child(1).attrs.sourceParagraph).toBeNull();
    expect(editor.state.doc.attrs.wordSectionState.breaks[0].paragraph).toBe(1);
    expect(wordSectionMap(editor.state)?.paragraphs.map((p) => p.sectionId)).toEqual([
      source.sections[0].id,
      source.sections[0].id,
      source.sections[1].id,
    ]);
    editor.commands.undo();
    expect(editor.getJSON()).toEqual(original);
  } finally {
    editor.destroy();
  }
});
it.each([1, 6, 10])(
  'moves the live break after Enter at position %s and restores it atomically',
  (pos) => {
    const editor = make();
    try {
      const initial = editor.state.doc.toJSON();
      editor.commands.setTextSelection(pos);
      editor.commands.splitBlock();
      expect(editor.state.doc.attrs.wordSectionState.breaks[0].paragraph).toBe(1);
      expect(wordSectionMap(editor.state)?.paragraphs.map((p) => p.sectionId)).toEqual([
        source.sections[0].id,
        source.sections[0].id,
        source.sections[1].id,
      ]);
      editor.commands.undo();
      expect(editor.state.doc.toJSON()).toEqual(initial);
      editor.commands.redo();
      expect(editor.state.doc.attrs.wordSectionState.breaks[0].paragraph).toBe(1);
    } finally {
      editor.destroy();
    }
  },
);
it('keeps the break with an empty paragraph when a heading changes type on Enter', () => {
  const editor = make('<h1 data-source-paragraph="0:0">AlphaBeta</h1>');
  try {
    editor.commands.setTextSelection(10);
    editor.commands.splitBlock();
    expect(editor.state.doc.child(1).type.name).toBe('paragraph');
    expect(editor.state.doc.attrs.wordSectionState.breaks[0].paragraph).toBe(1);
  } finally {
    editor.destroy();
  }
});
it.each(['joinBackward', 'joinForward'] as const)(
  'removes the break on %s and recovers it with undo',
  (command) => {
    const editor = make();
    try {
      const initial = editor.state.doc.toJSON();
      editor.commands.setTextSelection(command === 'joinBackward' ? 12 : 10);
      editor.commands[command]();
      expect(editor.getText()).toBe('AlphaBetaGamma');
      expect(editor.state.doc.attrs.wordSectionState.breaks).toEqual([]);
      expect(wordSectionMap(editor.state)?.ranges.map((r) => r.id)).toEqual([
        source.sections[1].id,
      ]);
      editor.commands.undo();
      expect(editor.state.doc.toJSON()).toEqual(initial);
    } finally {
      editor.destroy();
    }
  },
);
it('handles selection replacement and HTML paragraph paste without accepting copied section state', () => {
  const editor = make();
  try {
    editor.commands.setTextSelection({ from: 6, to: 14 });
    editor.commands.insertContent('X');
    expect(editor.state.doc.attrs.wordSectionState.breaks).toEqual([]);
    editor.commands.undo();
    editor.commands.setTextSelection(6);
    editor.commands.insertContent('<p>One</p><p>Two</p>');
    expect(
      Array.from(
        { length: editor.state.doc.childCount },
        (_, i) => editor.state.doc.child(i).textContent,
      ),
    ).toEqual(['Alpha', 'One', 'Two', 'Beta', 'Gamma']);
    expect(editor.state.doc.attrs.wordSectionState.breaks[0].paragraph).toBe(3);
    expect(editor.getHTML()).not.toContain('wordSectionState');
  } finally {
    editor.destroy();
  }
});
it('restores typed live boundaries with normalized HTML after a saved reload', () => {
  const editor = make();
  let restored: Editor | undefined;
  try {
    editor.commands.setTextSelection(6);
    editor.commands.splitBlock();
    const saved = wordSectionStateSchema.parse(editor.state.doc.attrs.wordSectionState);
    restored = new Editor({
      extensions: [...wordExtensions(), WordEditorSections.configure({ source })],
      content: editor.getHTML(),
    });
    updateWordSectionSource(restored, source, saved);
    expect(wordSectionMap(restored.state)).toEqual(wordSectionMap(editor.state));
    expect(restored.can().undo()).toBe(false);
    expect(() =>
      validateSectionState(
        { ...saved, breaks: [{ sectionId: 'unknown', paragraph: 0 }] },
        source,
        restored!.state.doc,
      ),
    ).toThrow();
    expect(initializeSectionState(restored.state.doc, source)).toBeNull();
  } finally {
    restored?.destroy();
    editor.destroy();
  }
});

it('tracks empty section paragraphs and rejects invalid saved boundaries before mutating the editor', () => {
  const editor = make('<p data-source-paragraph="0:0"></p>');
  try {
    editor.commands.setTextSelection(1);
    editor.commands.splitBlock();
    expect(editor.state.doc.attrs.wordSectionState.breaks[0].paragraph).toBe(1);
    editor.commands.undo();
    editor.commands.setTextSelection(3);
    editor.commands.joinBackward();
    expect(editor.state.doc.attrs.wordSectionState.breaks).toEqual([]);
    editor.commands.undo();
    editor.view.dispatch(
      editor.state.tr.setDocAttribute('wordSectionState', null).setMeta('addToHistory', false),
    );
    expect(() =>
      updateWordSectionSource(editor, source, {
        version: 1,
        finalSectionId: source.sections[1].id,
        breaks: [{ sectionId: source.sections[0].id, paragraph: 999 }],
      }),
    ).toThrow();
    expect(editor.state.doc.attrs.wordSectionState).toBeNull();
  } finally {
    editor.destroy();
  }
});
