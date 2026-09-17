import { it, expect } from 'vitest';
import { Editor } from '@tiptap/core';
import type { DocxStructure } from './docx-sections';
import { wordExtensions } from './word-extensions';
import {
  WordEditorSections,
  wordSectionMap,
  updateWordSectionSource,
} from './word-editor-sections';
import { showFormattingMarks } from './word-marks';

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
      drawings: [],
      notes: [],
    },
    {
      id: 'word/document.xml#section:1',
      endingParagraph: null,
      paragraphs: ['0:1'],
      propertiesXml: null,
      headers: [],
      footers: [],
      drawings: [],
      notes: [],
    },
  ],
};

it('hydrates source mapping without editing/history and caches positions across selection-only transactions', () => {
  const editor = new Editor({
    extensions: [...wordExtensions(), WordEditorSections],
    content: '<p data-source-paragraph="0:0">First</p><p data-source-paragraph="0:1">Last</p>',
  });
  try {
    const initialHTML = editor.getHTML();
    updateWordSectionSource(editor, source);
    const initial = wordSectionMap(editor.state);
    expect(initial?.ranges).toHaveLength(2);
    editor.commands.setTextSelection(3);
    expect(wordSectionMap(editor.state)).toBe(initial);
    showFormattingMarks(editor, true);
    expect(wordSectionMap(editor.state)).toBe(initial);
    expect(editor.view.dom.querySelectorAll('.section-mark')).toHaveLength(1);
    expect(editor.getHTML()).toBe(initialHTML);
    expect(editor.can().undo()).toBe(false);
    editor.commands.insertContent('more');
    expect(wordSectionMap(editor.state)?.ranges[0].to).toBe(11);
    editor.commands.undo();
    expect(wordSectionMap(editor.state)).toEqual(initial);
    updateWordSectionSource(editor, undefined);
    expect(editor.view.dom.querySelectorAll('.section-mark')).toHaveLength(0);
    expect(wordSectionMap(editor.state)?.ranges).toEqual([]);
  } finally {
    editor.destroy();
  }
});

it('removes ambiguous break decorations after a split and restores them with undo', () => {
  const editor = new Editor({
    extensions: [...wordExtensions(), WordEditorSections.configure({ source })],
    content: '<p data-source-paragraph="0:0">First</p><p data-source-paragraph="0:1">Last</p>',
  });
  try {
    showFormattingMarks(editor, true);
    editor.commands.setTextSelection(3);
    editor.commands.splitBlock();
    expect(wordSectionMap(editor.state)?.issues[0].code).toBe('duplicate-boundary');
    expect(editor.view.dom.querySelectorAll('.section-mark')).toHaveLength(0);
    editor.commands.undo();
    expect(editor.view.dom.querySelectorAll('.section-mark')).toHaveLength(1);
    expect(editor.getHTML()).not.toContain('Section break');
    showFormattingMarks(editor, false);
    expect(editor.view.dom.querySelectorAll('.section-mark')).toHaveLength(0);
  } finally {
    editor.destroy();
  }
});
