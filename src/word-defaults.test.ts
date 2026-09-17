import { expect, it, vi } from 'vitest';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { newFile } from './model';
import { wordExtensions } from './word-extensions';
import { exportOffice } from './formats';
import { readDocx, wordXml, descendants, val } from './docx-import';
import { wordDefaultFontBytes } from './word-default-font';

// Keep the real converter; adapt its Node input transport for this test runtime.
vi.mock('mammoth', async (original) => {
  const actual = await original<typeof import('mammoth')>();
  return {
    ...actual,
    convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
      actual.convertToHtml(
        { buffer: Buffer.from(input.arrayBuffer) },
        options as Parameters<typeof actual.convertToHtml>[1],
      ),
  };
});

for (const mode of ['empty', 'typed', 'legacy-unformatted'] as const)
  it(`exports and reimports the displayed normal paragraph defaults: ${mode}`, async () => {
    const file = newFile('word');
    if (file.content.kind !== 'word') throw Error();
    const editor = new Editor({ extensions: wordExtensions(), content: file.content.html });
    try {
      if (mode === 'typed') {
        editor.commands.insertContent('First');
        editor.commands.splitBlock();
        editor.commands.insertContent('Second');
      }
      file.content.html = mode === 'legacy-unformatted' ? '<p>Legacy</p>' : editor.getHTML();
      const blob = await exportOffice(file);
      const bytes = await new Promise<ArrayBuffer>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as ArrayBuffer);
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(blob);
      });
      const zip = await JSZip.loadAsync(bytes);
      const settings = wordXml(await zip.file('word/settings.xml')!.async('string'));
      const embedding = descendants(settings, 'embedTrueTypeFonts');
      expect(embedding).toHaveLength(1);
      expect(['', '1', 'true']).toContain(val(embedding[0]));
      const fontTable = wordXml(await zip.file('word/fontTable.xml')!.async('string'));
      const key = val(descendants(fontTable, 'embedRegular')[0], 'fontKey').replace(/[{}-]/g, '');
      expect(key).toMatch(/^[\da-f]{32}$/i);
      const mask = key
        .match(/../g)!
        .map((pair) => parseInt(pair, 16))
        .reverse();
      const embedded = await zip.file('word/fonts/font1.odttf')!.async('uint8array');
      for (let i = 0; i < 32; i++) embedded[i] ^= mask[i % 16];
      expect(embedded).toEqual(await wordDefaultFontBytes());
      const styles = wordXml(await zip.file('word/styles.xml')!.async('string'));
      const defaults = descendants(styles, 'docDefaults')[0];
      expect(val(descendants(defaults, 'rFonts')[0], 'ascii')).toBe('Inter');
      expect(val(descendants(defaults, 'sz')[0])).toBe('24');
      const spacing = descendants(defaults, 'spacing')[0];
      expect([val(spacing, 'line'), val(spacing, 'lineRule'), val(spacing, 'after')]).toEqual([
        '396',
        'auto',
        '240',
      ]);
      const imported = await readDocx(bytes);
      const restored = new Editor({ extensions: wordExtensions(), content: imported.content.html });
      try {
        const paragraphs = restored.getJSON().content!;
        expect(paragraphs).toHaveLength(mode === 'typed' ? 2 : 1);
        for (const p of paragraphs)
          expect(p.attrs).toMatchObject({
            paragraphFontFamily: 'Inter',
            paragraphFontSize: '12pt',
            paragraphLineHeight: '1.65',
            spaceAfter: '12pt',
          });
        expect(restored.getText()).toBe(
          mode === 'typed' ? 'First\n\nSecond' : mode === 'empty' ? '' : 'Legacy',
        );
      } finally {
        restored.destroy();
      }
    } finally {
      editor.destroy();
    }
  });
