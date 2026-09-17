import { Editor } from '@tiptap/core';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { newFile } from './model';
import { wordExtensions } from './word-extensions';
import { checkpointWordEditSession } from './word-edit-run';
import { exportOffice } from './formats';
import { descendants, val, wordXml } from './docx-import';

it('writes new DOCX save-session boundaries and a matching register, including inherited source text', async () => {
  for (const inherited of [false, true]) {
    const file = newFile('word');
    if (file.content.kind !== 'word') throw Error();
    const editor = new Editor({
      extensions: wordExtensions(),
      content: inherited ? '<p>before</p>' : file.content.html,
    });
    try {
      if (!inherited) editor.commands.insertContent('before');
      editor.commands.setTextSelection(editor.state.doc.content.size - 1);
      checkpointWordEditSession(editor.view);
      editor.commands.insertContent(' added');
      file.content.html = editor.getHTML();
      const blob = await exportOffice(file);
      const bytes = await new Promise<ArrayBuffer>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as ArrayBuffer);
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(blob);
      });
      const zip = await JSZip.loadAsync(bytes);
      const document = wordXml(await zip.file('word/document.xml')!.async('string'));
      const runs = descendants(document, 'r').filter((r) => descendants(r, 't').length);
      expect(
        runs.map((r) =>
          descendants(r, 't')
            .map((t) => t.textContent)
            .join(''),
        ),
      ).toEqual(['before', ' added']);
      const ids = runs.map((r) => val(r, 'rsidR'));
      expect(ids.every((id) => /^[0-9A-F]{8}$/.test(id))).toBe(true);
      expect(new Set(ids).size).toBe(2);
      const settings = wordXml(await zip.file('word/settings.xml')!.async('string'));
      const register = descendants(settings, 'rsids');
      expect(register).toHaveLength(1);
      expect(register[0].parentElement?.localName).toBe('settings');
      const registered = descendants(register[0], 'rsid').map((r) => val(r));
      expect(registered).toEqual(expect.arrayContaining(ids));
      expect(descendants(register[0], 'rsidRoot')).toHaveLength(1);
      expect(descendants(document, 'ins')).toHaveLength(0);
    } finally {
      editor.destroy();
    }
  }
});
