import { it, expect, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { wordExtensions, wordJSON } from './word-extensions';
import { wordSelectionLineSpacing } from './word-line-spacing-commands';
import { wordLineSpacing, type WordLineSpacing } from './word-line-spacing';
import { readDocx } from './docx-import';
import { exportOffice, importFile } from './formats';
import native from '../tests/fixtures/native-word-story-line-authoring.json';
vi.mock('mammoth', async original => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});
const spacing = (name: string): WordLineSpacing => name.endsWith('single') ? { rule: 'auto', line: 240 }
  : name.endsWith('double') ? { rule: 'auto', line: 480 }
  : name.endsWith('exact20') ? { rule: 'exact', line: 400 } : { rule: 'atLeast', line: 1000 };
for (const row of native.rows) it(`authors native ${row.name} semantics and actual retained DOCX`, async () => {
  const data = new Uint8Array(readFileSync(row.source)).buffer;
  const file = await importFile(Object.assign(new File([data], row.name + '.docx'), { arrayBuffer: async () => data }));
  if (file.content.kind !== 'word') throw Error('Word fixture');
  const reference = file.content.docxStructure!.sections[0][row.kind === 'header' ? 'headers' : 'footers'].find(r => r.type === 'default')!;
  const story = file.content.stories!.parts.find(p => p.kind === row.kind && p.relationshipIds.includes(reference.relationshipId))!;
  expect(wordJSON(story.html).content).toHaveLength(5);
  const editor = new Editor({ extensions: wordExtensions(), content: story.html });
  try {
    let pos = 1; editor.state.doc.forEach((node, offset, index) => { if (index === 2) pos = offset + 1; });
    editor.commands.setTextSelection(pos); const original = editor.getJSON();
    expect(editor.can().setWordLineSpacing(spacing(row.name))).toBe(true); expect(editor.can().undo()).toBe(false);
    editor.commands.setWordLineSpacing(spacing(row.name)); const changed = editor.getJSON();
    expect(wordSelectionLineSpacing(editor.state)).toEqual(spacing(row.name));
    editor.commands.setWordLineSpacing(spacing(row.name));
    editor.commands.undo(); expect(editor.getJSON()).toEqual(original); expect(editor.can().undo()).toBe(false);
    editor.commands.redo(); expect(editor.getJSON()).toEqual(changed);
    story.html = editor.getHTML();
    const zip = await JSZip.loadAsync(await exportOffice(file));
    const result = (await readDocx(await zip.generateAsync({ type: 'arraybuffer' }))).content;
    const output = wordJSON(result.stories!.parts.find(p => p.path === story.path)!.html).content!;
    expect(output).toHaveLength(5);
    for (let index = 0; index < 5; index++) {
      const attrs = output[index].attrs!;
      expect(wordLineSpacing(attrs.paragraphLineHeight, attrs.paragraphLineRule)).toEqual(index === 2 ? spacing(row.name) : { rule: 'exact', line: 800 });
    }
    const source = await JSZip.loadAsync(data);
    for (const path of ['word/document.xml', 'word/styles.xml']) expect(await zip.file(path)!.async('string')).toBe(await source.file(path)!.async('string'));
  } finally { editor.destroy(); }
});

it('shows mixed spacing and formats a range atomically without changing a following paragraph or other attributes', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p style="line-height:1;margin-top:8pt">A</p><p style="line-height:20pt;margin-bottom:6pt">B</p><p style="line-height:2">C</p>' });
  try {
    editor.commands.setTextSelection({ from: 1, to: 7 });
    expect(wordSelectionLineSpacing(editor.state)).toBeNull(); const before = editor.getJSON();
    editor.commands.setWordLineSpacing({ rule: 'atLeast', line: 1000 });
    expect(wordSelectionLineSpacing(editor.state)).toEqual({ rule: 'atLeast', line: 1000 });
    const changed = editor.getJSON().content!;
    expect(changed[0].attrs!.spaceBefore).toBe('8pt'); expect(changed[1].attrs!.spaceAfter).toBe('6pt');
    expect(changed[2]).toEqual(before.content![2]);
    editor.commands.undo(); expect(editor.getJSON()).toEqual(before); expect(editor.can().undo()).toBe(false);
  } finally { editor.destroy(); }
});

it('rejects malformed line spacing without changing content or history', () => {
  const editor = new Editor({ extensions: wordExtensions(), content: '<p style="line-height:20pt">Text</p>' });
  try {
    const before = editor.getJSON();
    for (const value of [null, {}, { rule: 'bad', line: 20 }, ...[0, -1, 0.5, 1, 13, 31681, NaN, Infinity].map(line => ({ rule: 'exact', line }))]) {
      expect(editor.commands.setWordLineSpacing(value as WordLineSpacing)).toBe(false);
      expect(editor.getJSON()).toEqual(before); expect(editor.can().undo()).toBe(false);
    }
  } finally { editor.destroy(); }
});
