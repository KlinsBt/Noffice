import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { readDocx, wordXml, descendants, val } from './docx-import';
import { wordExtensions } from './word-extensions';
import { resolveWordSections } from './word-section-layout';
import { exportOffice } from './formats';
import { newFile } from './model';
import native from '../tests/fixtures/native-word-script-prefix.json';

vi.mock('mammoth', async original => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});

for (const row of native.rows) it(`exports native scripted prefix typing with paragraph-terminal navigation: ${row.name}`, async () => {
  const bytes = readFileSync(`tests/fixtures/word-script-scopes/${row.name}-formatted.docx`);
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(row.sourceHash);
  const data = new Uint8Array(bytes).buffer, { content } = await readDocx(data);
  const section = resolveWordSections(content)[0];
  const relationship = section[row.kind === 'Headers' ? 'headers' : 'footers'].default;
  const part = row.kind === 'Body' ? undefined : content.stories!.parts.find(p => p.relationshipIds.includes(relationship!.relationshipId))!;
  const editor = new Editor({ extensions: wordExtensions({ retainedEditRuns: true }), content: part?.html || content.html });
  try {
    editor.commands.setTextSelection(1); editor.view.dispatch(editor.state.tr.insertText('Q'));
    if (part) part.html = editor.getHTML(); else content.html = editor.getHTML();
    const file = { ...newFile('word', row.name), content,
      original: { data, name: row.name + '.docx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', contentFingerprint: 'edited' } };
    const output = await JSZip.loadAsync(await exportOffice(file));
    const input = await JSZip.loadAsync(data), changedPart = part?.path || 'word/document.xml';
    expect(Object.keys(output.files).sort()).toEqual(Object.keys(input.files).sort());
    for (const path of Object.keys(input.files)) if (!input.files[path].dir && ![changedPart, 'word/settings.xml'].includes(path))
      expect(await output.file(path)!.async('uint8array'), path).toEqual(await input.file(path)!.async('uint8array'));
    const restored = (await readDocx(await output.generateAsync({ type: 'arraybuffer' }))).content;
    editor.commands.setContent(part ? restored.stories!.parts.find(p => p.path === part.path)!.html : restored.html);
    const p = editor.getJSON().content![0];
    const characters = [...p.content!.flatMap(n => [...(n.type === 'wordTab' ? '\t' : 'text' in n ? n.text : '')].map(text => ({
      text, superscript: n.marks?.some(m => m.type === 'superscript') ? -1 : 0,
      subscript: n.marks?.some(m => m.type === 'subscript') ? -1 : 0,
      size: parseFloat(String(n.marks?.find(m => m.type === 'textStyle')?.attrs?.fontSize || p.attrs!.paragraphFontSize)),
    }))), { text: '\r', superscript: p.attrs!.paragraphScript === 'superscript' ? -1 : 0,
      subscript: p.attrs!.paragraphScript === 'subscript' ? -1 : 0, size: parseFloat(p.attrs!.paragraphFontSize) }];
    expect(characters).toEqual(row.reopened.map(({ text, superscript, subscript, size }) => ({ text, superscript, subscript, size })));
    if (row.kind === 'Body') {
      const xml = wordXml(await output.file(changedPart)!.async('string'));
      const start = descendants(xml, 'bookmarkStart').find(e => val(e, 'name') === '_GoBack')!;
      expect(start.previousElementSibling?.textContent).toBe('Q');
      expect(start.nextElementSibling?.localName).toBe('bookmarkEnd');
      expect(val(start.nextElementSibling!, 'id')).toBe(val(start, 'id'));
      expect(descendants(xml, 'bookmarkEnd')).toHaveLength(1);
    }
  } finally { editor.destroy(); }
});
