import { describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { Document, Packer, Paragraph } from 'docx';
import { readDocx, WORD_NS, descendants, wordXml, val } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { newFile } from './model';
import { wordJSON, wordExtensions } from './word-extensions';
import { Editor } from '@tiptap/core';
import { WordEditorSections, updateWordSectionSource } from './word-editor-sections';
import { exportOffice } from './formats';
import { hydrateWordStructure } from './word-structure';
import { contentFingerprint } from './office-preservation';

// Mammoth's Node transport expects Buffer; its browser transport expects ArrayBuffer.
// Exercise the real converter with the appropriate test-runtime transport.
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
async function setup(
  body = '<w:p><w:r><w:rPr><w:b/><w:lang w:val="de-DE"/></w:rPr><w:t>Hello </w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="185ABD"/></w:rPr><w:t>world</w:t></w:r></w:p>',
) {
  const data = await Packer.toArrayBuffer(
    new Document({ sections: [{ children: [new Paragraph('Initial')] }] }),
  );
  const zip = await JSZip.loadAsync(data);
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${WORD_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}<w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1000" w:bottom="1200" w:left="1400" w:right="1600"/></w:sectPr></w:body></w:document>`,
  );
  zip.file(
    'word/header1.xml',
    `<w:hdr xmlns:w="${WORD_NS}"><w:p><w:r><w:t>KEEP HEADER</w:t></w:r></w:p></w:hdr>`,
  );
  zip.file('customXml/item1.xml', '<retained custom="yes"/>');
  const relationships = await zip.file('word/_rels/document.xml.rels')!.async('string');
  zip.file(
    'word/_rels/document.xml.rels',
    relationships.replace(
      '</Relationships>',
      ['rHeader', 'rInherited']
        .map(
          (id) =>
            `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`,
        )
        .join('') + '</Relationships>',
    ),
  );
  const source = await zip.generateAsync({ type: 'arraybuffer' });
  const read = await readDocx(source);
  const file = newFile('word', 'Retained', read.content);
  file.original = { name: 'source.docx', data: source };
  return { file, zip, read };
}
async function exported(file: Awaited<ReturnType<typeof setup>>['file']) {
  const blob = await exportRetainedDocument(file);
  const bytes = await new Promise<ArrayBuffer>((resolve) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as ArrayBuffer);
    r.readAsArrayBuffer(blob);
  });
  const zip = await JSZip.loadAsync(bytes);
  const doc = wordXml(await zip.file('word/document.xml')!.async('string'));
  return { zip, doc, bytes };
}
function html(file: Awaited<ReturnType<typeof setup>>['file']) {
  if (file.content.kind !== 'word') throw Error();
  const dom = new DOMParser().parseFromString(file.content.html, 'text/html');
  return {
    dom,
    save: () => {
      if (file.content.kind === 'word') file.content.html = dom.body.innerHTML;
    },
  };
}
describe('retained Word editing', () => {
  it('restores RGB from an unchanged legacy snapshot and retains its original identity', async () => {
    const { file } = await setup();
    file.content = (await readDocx(file.original!.data, { legacyRunColors: true })).content;
    file.original!.contentFingerprint = await contentFingerprint(file.content);
    const original = new Uint8Array(file.original!.data).slice();
    const migrated = await hydrateWordStructure(file);
    expect(migrated.content.kind).toBe('word');
    if (migrated.content.kind !== 'word') throw Error();
    expect(migrated.content.runColorsVersion).toBe(1);
    expect(migrated.content.html).toContain('rgb(24, 90, 189)');
    expect(migrated.original!.contentFingerprint).toBe(await contentFingerprint(migrated.content));
    expect(new Uint8Array(migrated.original!.data)).toEqual(original);
    expect(await hydrateWordStructure(migrated)).toBe(migrated);
  });
  it('keeps legacy edits and source RGB when their missing color cannot be safely migrated', async () => {
    const { file } = await setup();
    file.content = (await readDocx(file.original!.data, { legacyRunColors: true })).content;
    file.original!.contentFingerprint = await contentFingerprint(file.content);
    const edit = html(file);
    edit.dom.querySelector('em')!.textContent = 'legacy edit';
    edit.save();
    const editedHtml = file.content.html;
    const migrated = await hydrateWordStructure(file);
    if (migrated.content.kind !== 'word') throw Error();
    expect(migrated.content.html).toBe(editedHtml);
    expect(migrated.content.runColorsVersion).toBeUndefined();
    const { doc } = await exported(migrated);
    expect(descendants(doc, 't').map((node) => node.textContent).join('')).toBe('Hello legacy edit');
    expect(val(descendants(doc, 'color')[0])).toBe('185ABD');
  });
  it.each(['left', 'bar'])('edits text without counting %s stop definitions as literal tabs', async (alignment) => {
    const { file } = await setup(`<w:p><w:pPr><w:tabs><w:tab w:val="${alignment}" w:pos="1440"/></w:tabs></w:pPr><w:r><w:t>Before</w:t><w:tab/><w:t>After</w:t></w:r></w:p>`);
    const original = new Uint8Array(file.original!.data).slice();
    const edit = html(file);
    const walker = edit.dom.createTreeWalker(edit.dom.body, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) walker.currentNode.textContent = walker.currentNode.textContent!.replace('Before', 'Edited');
    edit.save();
    const result = await exported(file);
    const paragraph = descendants(result.doc, 'p')[0];
    expect(descendants(paragraph, 't').map((node) => node.textContent).join('')).toBe('EditedAfter');
    expect(descendants(paragraph, 'tabs')[0].outerHTML).toContain(`w:val="${alignment}"`);
    expect(descendants(paragraph, 'tab')).toHaveLength(2);
    expect(new Uint8Array(file.original!.data)).toEqual(original);
  });
  it.each([
    [
      'named bookmark',
      '<w:bookmarkStart w:id="7" w:name="UserPlace"/><w:bookmarkEnd w:id="7"/><w:r><w:t>Delete</w:t></w:r>',
    ],
    [
      'noncollapsed navigation bookmark',
      '<w:bookmarkStart w:id="7" w:name="_GoBack"/><w:r><w:t>Delete</w:t></w:r><w:bookmarkEnd w:id="7"/>',
    ],
    [
      'duplicate navigation identity',
      '<w:bookmarkStart w:id="7" w:name="_GoBack"/><w:bookmarkEnd w:id="7"/><w:bookmarkStart w:id="7" w:name="_GoBack"/><w:bookmarkEnd w:id="7"/><w:r><w:t>Delete</w:t></w:r>',
    ],
    ['field', '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>'],
    ['foreign payload', '<x:r xmlns:x="urn:unknown"><w:r><w:t>Delete</w:t></w:r></x:r>'],
    ['unknown script payload', '<w:pPr><w:rPr><w:vertAlign w:val="superscript" future="retained"/></w:rPr></w:pPr><w:r><w:t>Delete</w:t></w:r>'],
  ])('rejects whole-paragraph deletion of %s without changing source bytes', async (_name, xml) => {
    const { file } = await setup(`<w:p>${xml}</w:p><w:p><w:r><w:t>Keep</w:t></w:r></w:p>`);
    const original = new Uint8Array(file.original!.data).slice();
    const edit = html(file);
    edit.dom.querySelector('p')!.remove();
    edit.save();
    await expect(exported(file)).rejects.toThrow('document structures');
    expect(new Uint8Array(file.original!.data)).toEqual(original);
  });

  it('deletes modeled line breaks and only a collapsed native navigation bookmark', async () => {
    const { file } = await setup(
      '<w:p><w:pPr><w:widowControl/></w:pPr><w:bookmarkStart w:id="7" w:name="_GoBack"/><w:bookmarkEnd w:id="7"/><w:r><w:rPr><w:kern w:val="16"/></w:rPr><w:t>Delete</w:t><w:br/><w:lastRenderedPageBreak/><w:t>also</w:t></w:r></w:p><w:p><w:r><w:t>Keep</w:t></w:r></w:p>',
    );
    const edit = html(file);
    edit.dom.querySelector('p')!.remove();
    edit.save();
    const { doc } = await exported(file);
    expect(descendants(doc, 't').map((t) => t.textContent)).toEqual(['Keep']);
    expect(descendants(doc, 'bookmarkStart')).toHaveLength(0);
  });
  it.each(['baseline', 'superscript', 'subscript'])('deletes a modeled %s paragraph mark without changing retained originals', async script => {
    const { file } = await setup(`<w:p><w:pPr><w:rPr><w:vertAlign w:val="${script}"/></w:rPr></w:pPr><w:r><w:t>Delete</w:t></w:r></w:p><w:p><w:r><w:t>Keep</w:t></w:r></w:p>`);
    const original = file.original!.data.slice(0), edit = html(file);
    edit.dom.querySelector('p')!.remove(); edit.save();
    expect(descendants((await exported(file)).doc, 't').map(t => t.textContent)).toEqual(['Keep']);
    expect(file.original!.data).toEqual(original);
  });

  it('exports settled column run boundaries and rejects a stale content binding', async () => {
    const data = Uint8Array.from(
      readFileSync('tests/fixtures/word-unequal-flow-narrow-first-wrapped.docx'),
    ).buffer;
    const read = await readDocx(data),
      file = newFile('word', 'Wrapped export', read.content);
    file.original = { name: 'wrapped.docx', data };
    if (file.content.kind !== 'word') throw Error('Expected Word');
    const editor = new Editor({ extensions: wordExtensions(), content: file.content.html });
    const originalText = editor.state.doc.textContent;
    try {
      editor.commands.setTextSelection(editor.state.doc.content.size - 1);
      editor.commands.setHardBreak();
      editor.commands.insertContent('added');
      file.content.html = editor.getHTML();
      const layout = {
        content: JSON.stringify(file.content),
        paragraphs: [{ text: originalText + '\nadded', boundaries: [140, 420] }],
      };
      const blob = await exportRetainedDocument(file, layout);
      const bytes = await new Promise<ArrayBuffer>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as ArrayBuffer);
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(blob);
      });
      const zip = await JSZip.loadAsync(bytes);
      const doc = wordXml(await zip.file('word/document.xml')!.async('string'));
      expect(descendants(doc, 't').map((t) => t.textContent?.length)).toEqual([140, 280, 139, 5]);
      expect(descendants(doc, 'br')).toHaveLength(1);
      const breakRun = descendants(doc, 'br')[0].parentElement!;
      expect(descendants(breakRun, 'rFonts')[0].getAttributeNS(WORD_NS, 'ascii')).toBe('Arial');
      expect(descendants(breakRun, 'sz')[0].getAttributeNS(WORD_NS, 'val')).toBe('20');
      expect(descendants(doc, 'lastRenderedPageBreak')).toHaveLength(0);
      expect(new Uint8Array(file.original.data)).toEqual(new Uint8Array(data));
      file.content.orientation = 'landscape';
      await expect(exportRetainedDocument(file, layout)).rejects.toThrow('changed before export');
    } finally {
      editor.destroy();
    }
  });
  it('keeps the native wrapped-column source package after an editor round trip and restored text', async () => {
    const data = Uint8Array.from(
      readFileSync('tests/fixtures/word-unequal-flow-narrow-first-wrapped.docx'),
    ).buffer;
    const read = await readDocx(data),
      file = newFile('word', 'Wrapped source', read.content);
    file.original = { name: 'wrapped.docx', data };
    if (file.content.kind !== 'word') throw Error('Expected Word');
    const editor = new Editor({ extensions: wordExtensions(), content: file.content.html });
    try {
      editor.commands.setTextSelection(editor.state.doc.content.size - 1);
      editor.commands.setHardBreak();
      editor.commands.insertContent('added');
      editor.commands.deleteRange({
        from: editor.state.doc.content.size - 7,
        to: editor.state.doc.content.size - 1,
      });
      file.content.html = editor.getHTML();
      expect(editor.state.doc.textContent).toBe(
        wordJSON(read.content.html)
          .content![0].content!.map((n) => n.text || '')
          .join(''),
      );
      expect(new Uint8Array((await exported(file)).bytes)).toEqual(new Uint8Array(data));
    } finally {
      editor.destroy();
    }
  });
  it('invalidates an edited paragraph pagination cache while retaining explicit breaks and the original', async () => {
    const { file } = await setup(
      '<w:p><w:r><w:t>First</w:t><w:lastRenderedPageBreak/><w:br w:type="page"/><w:t>Second</w:t></w:r></w:p>',
    );
    expect(new Uint8Array((await exported(file)).bytes)).toEqual(
      new Uint8Array(file.original!.data),
    );
    const edit = html(file);
    edit.dom.querySelector('p')!.append(edit.dom.createElement('br'), 'Third');
    edit.save();
    const result = await exported(file);
    expect(descendants(result.doc, 'lastRenderedPageBreak')).toHaveLength(0);
    expect(descendants(result.doc, 'br').map((e) => val(e, 'type'))).toEqual(['page', '']);
    const original = await JSZip.loadAsync(file.original!.data);
    expect(await original.file('word/document.xml')!.async('string')).toContain(
      'lastRenderedPageBreak',
    );
  });
  it('preserves a leading collapsed bookmark while adding a later hard line break', async () => {
    const { file } = await setup(
      '<w:p><w:bookmarkStart w:id="7" w:name="Position"/><w:bookmarkEnd w:id="7"/><w:r><w:t>First</w:t><w:br/><w:t>Second</w:t></w:r></w:p>',
    );
    const edit = html(file);
    edit.dom.querySelector('p')!.append(edit.dom.createElement('br'), 'Third');
    edit.save();
    const result = await exported(file);
    const p = descendants(result.doc, 'p')[0];
    expect(
      [...p.children]
        .filter((e) => e.localName !== 'pPr')
        .slice(0, 2)
        .map((e) => [e.localName, val(e, 'id')]),
    ).toEqual([
      ['bookmarkStart', '7'],
      ['bookmarkEnd', '7'],
    ]);
    expect(descendants(p, 'br')).toHaveLength(2);
    expect(
      descendants(p, 't')
        .map((e) => e.textContent)
        .join(''),
    ).toBe('FirstSecondThird');
  });
  it.each(['page', 'column'] as const)(
    'keeps a named collapsed anchor before a leading native %s command through export and history',
    async (kind) => {
      const { file } = await setup(
        '<w:p><w:bookmarkStart w:id="7" w:name="Position"/><w:bookmarkEnd w:id="7"/><w:r><w:t>Before</w:t></w:r></w:p>',
      );
      if (file.content.kind !== 'word') throw Error('Expected Word');
      const content = file.content;
      const editor = new Editor({
        extensions: [...wordExtensions(), WordEditorSections],
        content: content.html,
      });
      try {
        updateWordSectionSource(editor, content.docxStructure);
        editor.commands.setTextSelection(1);
        const initial = editor.getJSON();
        editor.commands.insertWordFlowBreak(kind);
        const inserted = editor.getJSON();
        editor.commands.undo();
        expect(editor.getJSON()).toEqual(initial);
        editor.commands.redo();
        expect(editor.getJSON()).toEqual(inserted);
        content.html = editor.getHTML();
        content.sectionState = editor.state.doc.attrs.wordSectionState;
        const result = await exported(file);
        const paragraphs = descendants(result.doc, 'p');
        expect(paragraphs).toHaveLength(kind === 'page' ? 2 : 1);
        expect(
          [...paragraphs[0].children]
            .filter((e) => e.localName !== 'pPr')
            .slice(0, 2)
            .map((e) => [e.localName, val(e, 'id')]),
        ).toEqual([
          ['bookmarkStart', '7'],
          ['bookmarkEnd', '7'],
        ]);
        expect(descendants(paragraphs[0], 'br').map((e) => val(e, 'type'))).toEqual([kind]);
        expect(
          descendants(result.doc, 't')
            .map((e) => e.textContent)
            .join(''),
        ).toBe('Before');
      } finally {
        editor.destroy();
      }
    },
  );
  it('still rejects hard-line edits across a nonempty bookmark', async () => {
    const { file } = await setup(
      '<w:p><w:bookmarkStart w:id="7" w:name="Range"/><w:r><w:t>First</w:t><w:br/><w:t>Second</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>',
    );
    const edit = html(file);
    edit.dom.querySelector('p')!.append(edit.dom.createElement('br'), 'Third');
    edit.save();
    await expect(exported(file)).rejects.toThrow('complex paragraph structures');
  });
  it.each([false, true])('preserves native prefix affinity in a tabbed paragraph with named anchor %s', async (named) => {
    const { file } = await setup('<w:p>' +
      (named ? '<w:bookmarkStart w:id="8" w:name="NofficeAnchor"/>' : '') +
      '<w:bookmarkStart w:id="7" w:name="_GoBack"/>' +
      (named ? '<w:bookmarkEnd w:id="8"/>' : '') +
      '<w:bookmarkEnd w:id="7"/><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:tab/><w:t>C</w:t></w:r></w:p>');
    const original = file.original!.data.slice(0);
    const edit = html(file); edit.dom.querySelector('p')!.prepend('Saved '); edit.save();
    const result = await exported(file), p = descendants(result.doc, 'p')[0];
    let offset = 0;
    const positions: [string, string, number][] = [];
    for (const e of descendants(p, '*')) {
      if (e.localName === 't') offset += e.textContent!.length;
      else if (e.localName === 'tab') offset++;
      else if (['bookmarkStart', 'bookmarkEnd'].includes(e.localName)) positions.push([e.localName, val(e, 'id'), offset]);
    }
    expect(positions).toEqual([
      ...(named ? [['bookmarkStart', '8', 0], ['bookmarkEnd', '8', 0]] : []),
      ['bookmarkStart', '7', 6], ['bookmarkEnd', '7', 6],
    ]);
    expect(wordJSON((await readDocx(result.bytes)).content.html).content![0].content!.map((n) => n.type === 'wordTab' ? '\t' : n.text || '').join('')).toBe('Saved A\tB\tC');
    expect(file.original!.data).toEqual(original);
  });
  it.each([
    '<w:bookmarkStart w:id="7" w:name="_GoBack"/><w:bookmarkEnd w:id="7"/><w:bookmarkStart w:id="7" w:name="Other"/><w:bookmarkEnd w:id="7"/>',
    '<w:bookmarkStart w:id="7" w:name="_GoBack"/><w:bookmarkEnd w:id="8"/>',
    '<w:bookmarkStart w:id="7" w:name="_Unknown"/><w:bookmarkEnd w:id="7"/>',
    '<w:bookmarkStart w:id="7" w:name="Range"/>',
  ])('rejects unmapped prefix bookmark identities: %s', async (prefix) => {
    const { file } = await setup('<w:p>' + prefix + '<w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r></w:p>');
    const edit = html(file); edit.dom.querySelector('p')!.prepend('Saved '); edit.save();
    await expect(exported(file)).rejects.toThrow('complex paragraph structures');
  });
  it.each([
    ['named range', 'w:id="7" w:name="UserPlace"', '<w:bookmarkEnd w:id="7"/>'],
    ['unknown start attribute', 'w:id="7" w:name="_GoBack" w:colFirst="1"', '<w:bookmarkEnd w:id="7"/>'],
    ['unknown end attribute', 'w:id="7" w:name="_GoBack"', '<w:bookmarkEnd w:id="7" w:unknown="1"/>'],
    ['missing end', 'w:id="7" w:name="_GoBack"', ''],
    ['mismatched end', 'w:id="7" w:name="_GoBack"', '<w:bookmarkEnd w:id="8"/>'],
    ['crossed paragraph', 'w:id="7" w:name="_GoBack"', '<w:p><w:r><w:t>Other</w:t></w:r></w:p><w:bookmarkEnd w:id="7"/>'],
    ['duplicate identity', 'w:id="7" w:name="_GoBack"', '<w:bookmarkEnd w:id="7"/><w:p><w:bookmarkStart w:id="7" w:name="Other"/><w:bookmarkEnd w:id="7"/></w:p>'],
    ['duplicate navigation name', 'w:id="7" w:name="_GoBack"', '<w:bookmarkEnd w:id="7"/><w:p><w:bookmarkStart w:id="8" w:name="_GoBack"/><w:bookmarkEnd w:id="8"/></w:p>'],
  ])('preserves rejected paragraph-terminal navigation: %s', async (_name, attributes, suffix) => {
    const { file } = await setup(`<w:p><w:bookmarkStart ${attributes}/><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t></w:r></w:p>${suffix}`);
    const edit = html(file); edit.dom.querySelector('p')!.prepend('Q'); edit.save();
    const before = structuredClone(file);
    await expect(exported(file)).rejects.toThrow('complex paragraph structures');
    expect(file).toEqual(before);
  });
  it('retains the mixed-wrap original and unrelated parts while matching native plain-run merging after editing', async () => {
    const input = readFileSync('tests/fixtures/word-mixed-wrap.docx');
    const data = Uint8Array.from(input).buffer;
    const source = await JSZip.loadAsync(data);
    const read = await readDocx(data);
    const file = newFile('word', 'Mixed wrap', read.content);
    file.original = { name: 'mixed.docx', data };
    expect(new Uint8Array((await exported(file)).bytes)).toEqual(new Uint8Array(data));
    if (file.content.kind !== 'word') throw new Error('Expected Word content');
    file.content.html = file.content.html.replace('Start', 'Edited');
    const output = await exported(file);
    expect(descendants(output.doc, 'p').map((p) => descendants(p, 'r').length)).toEqual([
      1, 13, 13, 1,
    ]);
    for (const [path, entry] of Object.entries(source.files)) {
      if (entry.dir || path === 'word/document.xml') continue;
      expect(await output.zip.file(path)!.async('uint8array'), path).toEqual(
        await entry.async('uint8array'),
      );
    }
    expect(new Uint8Array(file.original.data)).toEqual(new Uint8Array(data));
  });
  it('imports the pinned Word fallback size and preserves untouched native formatting', async () => {
    const source = readFileSync('tests/fixtures/word-line-metrics.docx');
    const bytes = new Uint8Array(source).buffer;
    const read = await readDocx(bytes);
    const doc = new DOMParser().parseFromString(read.content.html, 'text/html');
    expect(doc.querySelector('p')!.style.fontSize).toBe('10pt');
    const json = wordJSON(read.content.html);
    expect(json.content![0].attrs).toMatchObject({
      paragraphFontSize: '10pt',
      paragraphFontFamily: 'Arial',
    });
    const small = json.content![0].content!.find((n) => n.text === 'Small line')!;
    expect(small.marks!.find((m) => m.type === 'textStyle')!.attrs!.fontSize).toBe('10pt');
    const large = json.content![0].content!.find((n) => n.text === 'Large line')!;
    expect(large.marks!.find((m) => m.type === 'textStyle')!.attrs!.fontSize).toBe('20pt');
    const file = newFile('word', 'Fallback fonts', read.content);
    file.original = { name: 'fallback.docx', data: bytes };
    const result = await exported(file);
    const original = await JSZip.loadAsync(bytes);
    expect(await result.zip.file('word/document.xml')!.async('string')).toBe(
      await original.file('word/document.xml')!.async('string'),
    );
  });

  it('migrates omitted legacy fonts without replacing edits, explicit runs or source bytes', async () => {
    const bytes = new Uint8Array(readFileSync('tests/fixtures/word-line-metrics.docx')).buffer;
    const read = await readDocx(bytes);
    const expected = wordJSON(read.content.html);
    const file = newFile('word', 'Legacy fonts', read.content);
    file.original = { name: 'source.docx', data: bytes };
    if (file.content.kind !== 'word') throw Error();
    delete file.content.fontMetricsVersion;
    const h = html(file);
    for (const p of h.dom.querySelectorAll('p')) {
      p.style.removeProperty('font-size');
      p.style.removeProperty('font-family');
    }
    for (const span of h.dom.querySelectorAll('span'))
      if (span.style.fontSize === '10pt') span.style.removeProperty('font-size');
    h.save();
    file.original.contentFingerprint = await contentFingerprint(file.content);
    const migrated = await hydrateWordStructure(file);
    expect(migrated.revision).toBe(file.revision);
    expect(migrated.original!.data).toBe(bytes);
    expect(migrated.original!.contentFingerprint).toBe(await contentFingerprint(migrated.content));
    if (migrated.content.kind !== 'word') throw Error();
    expect(wordJSON(migrated.content.html)).toEqual(expected);
    expect(await hydrateWordStructure(migrated)).toBe(migrated);
    h.dom.querySelectorAll('p')[4].querySelector('span')!.textContent += ' saved edit';
    h.save();
    const edited = await hydrateWordStructure(file);
    expect(edited.original).toBe(file.original);
    if (edited.content.kind !== 'word') throw Error();
    expect(edited.content.html).toContain('saved edit');
    expect(edited.content.html).toContain('20pt');
    const exportedLegacy = await exported(file);
    expect(
      descendants(exportedLegacy.doc, 't')
        .map((t) => t.textContent)
        .join(''),
    ).toContain('saved edit');
    // Recovery supplies effective fonts without writing them onto untouched source runs.
    const ps = descendants(exportedLegacy.doc, 'p');
    expect(descendants(ps[0], 'sz').map((el) => val(el))).toEqual(['40']);
  });

  it('retains paragraph-mark fonts separately from text and writes them for new DOCX files', async () => {
    const { file } = await setup(
      '<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="48"/></w:rPr></w:pPr><w:r><w:t>Small text</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:sz w:val="36"/></w:rPr></w:pPr></w:p>',
    );
    if (file.content.kind !== 'word') throw Error();
    const json = wordJSON(file.content.html);
    expect(json.content![0].attrs!.paragraphFontSize).toBe('24pt');
    expect(
      json.content![0].content![0].marks!.find((m) => m.type === 'textStyle')!.attrs!.fontSize,
    ).toBe('10pt');
    expect(json.content![1].attrs!.paragraphFontSize).toBe('18pt');
    delete file.original;
    // This branch authors a body-only document from the inspected paragraphs.
    // Retained header creation is tested separately rather than silently dropped.
    delete file.content.stories;
    const blob = await exportOffice(file);
    const bytes = await new Promise<ArrayBuffer>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.readAsArrayBuffer(blob);
    });
    const reread = await readDocx(bytes);
    const after = wordJSON(reread.content.html);
    expect(after.content![0].attrs!.paragraphFontSize).toBe('24pt');
    expect(after.content![1].attrs!.paragraphFontSize).toBe('18pt');
    expect(
      after.content![0].content![0].marks!.find((m) => m.type === 'textStyle')!.attrs!.fontSize,
    ).toBe('10pt');
  });

  it('imports omitted native line spacing as single without materializing an export change', async () => {
    const { file } = await setup('<w:p><w:r><w:t>Default single</w:t></w:r></w:p>');
    if (file.content.kind !== 'word') throw Error();
    expect(wordJSON(file.content.html).content![0].attrs).toMatchObject({
      paragraphLineHeight: '1',
      paragraphLineRule: null,
    });
    const result = await exported(file);
    expect(descendants(result.doc, 'spacing')).toHaveLength(0);
  });
  it('hydrates an omitted legacy minimum without replacing edits or an explicit spacing choice', async () => {
    const { file } = await setup(
      '<w:p><w:pPr><w:spacing w:line="240" w:lineRule="atLeast"/></w:pPr><w:r><w:t>Minimum line</w:t></w:r></w:p>',
    );
    if (file.content.kind !== 'word') throw Error();
    delete file.content.lineSpacingVersion;
    const h = html(file),
      p = h.dom.querySelector('p')!;
    p.style.removeProperty('line-height');
    p.removeAttribute('data-word-line-rule');
    p.append(' older edit');
    h.save();
    const upgraded = await hydrateWordStructure(file);
    expect(upgraded.revision).toBe(file.revision);
    expect(upgraded.original).toBe(file.original);
    if (upgraded.content.kind !== 'word') throw Error();
    expect(upgraded.content.html).toContain('older edit');
    expect(wordJSON(upgraded.content.html).content![0].attrs).toMatchObject({
      paragraphLineHeight: '12pt',
      paragraphLineRule: 'atLeast',
    });
    const result = await exported(file);
    expect(val(descendants(result.doc, 'spacing')[0], 'lineRule')).toBe('atLeast');
    expect(
      descendants(result.doc, 't')
        .map((t) => t.textContent)
        .join(''),
    ).toContain('older edit');
    p.style.lineHeight = '1.5';
    h.save();
    const explicit = await hydrateWordStructure(file);
    if (explicit.content.kind !== 'word') throw Error();
    expect(wordJSON(explicit.content.html).content![0].attrs).toMatchObject({
      paragraphLineHeight: '1.5',
      paragraphLineRule: null,
    });
    p.style.removeProperty('line-height');
    h.save();
    file.original!.contentFingerprint = await contentFingerprint(file.content);
    const unchanged = await hydrateWordStructure(file);
    expect(unchanged.original!.contentFingerprint).toBe(
      await contentFingerprint(unchanged.content),
    );
  });
  it('imports minimum spacing and preserves its rule through normalization and edits', async () => {
    const { file } = await setup(
      '<w:p><w:pPr><w:spacing w:line="240" w:lineRule="atLeast"/></w:pPr><w:r><w:t>Minimum line</w:t></w:r></w:p>',
    );
    const h = html(file),
      p = h.dom.querySelector('p')!;
    expect(p.style.lineHeight).toBe('12pt');
    expect(p.getAttribute('data-word-line-rule')).toBe('atLeast');
    const attrs = wordJSON(h.dom.body.innerHTML).content![0].attrs!;
    expect(attrs.paragraphLineRule).toBe('atLeast');
    p.append(' edited');
    p.style.lineHeight = '18pt';
    h.save();
    let result = await exported(file);
    expect(
      descendants(result.doc, 'spacing').map((s) => [val(s, 'line'), val(s, 'lineRule')]),
    ).toEqual([['360', 'atLeast']]);
    p.removeAttribute('data-word-line-rule');
    h.save();
    result = await exported(file);
    expect(val(descendants(result.doc, 'spacing')[0], 'lineRule')).toBe('exact');
  });
  it('writes exact physical line spacing into a newly authored DOCX', async () => {
    const file = newFile('word');
    if (file.content.kind !== 'word') throw Error();
    file.content.html =
      '<p style="line-height:18pt">Point spacing</p><p style="line-height:24px">Pixel spacing</p><p style="line-height:18pt" data-word-line-rule="atLeast">Minimum spacing</p>';
    const blob = await exportOffice(file);
    const bytes = await new Promise<ArrayBuffer>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.readAsArrayBuffer(blob);
    });
    const zip = await JSZip.loadAsync(bytes),
      xml = wordXml(await zip.file('word/document.xml')!.async('string'));
    expect(descendants(xml, 'spacing').map((s) => [val(s, 'line'), val(s, 'lineRule')])).toEqual([
      ['360', 'exact'],
      ['360', 'exact'],
      ['360', 'atLeast'],
    ]);
  });
  it('imports exact line spacing and preserves or explicitly changes its native rule', async () => {
    const { file } = await setup(
      '<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>Exact line</w:t></w:r></w:p>',
    );
    const h = html(file),
      p = h.dom.querySelector('p')!;
    expect(p.style.lineHeight).toBe('12pt');
    p.append(' edited');
    h.save();
    let result = await exported(file);
    expect(val(descendants(result.doc, 'spacing')[0], 'lineRule')).toBe('exact');
    expect(val(descendants(result.doc, 'spacing')[0], 'line')).toBe('240');
    p.style.lineHeight = '18pt';
    h.save();
    result = await exported(file);
    expect(val(descendants(result.doc, 'spacing')[0], 'lineRule')).toBe('exact');
    expect(val(descendants(result.doc, 'spacing')[0], 'line')).toBe('360');
    p.style.lineHeight = '1.5';
    h.save();
    result = await exported(file);
    expect(val(descendants(result.doc, 'spacing')[0], 'lineRule')).toBe('auto');
    expect(val(descendants(result.doc, 'spacing')[0], 'line')).toBe('360');
  });
  it.each([false, true])('keeps unknown font-feature payload guarded during paragraph joins: %s', async (unknown) => {
    const ns = 'http://schemas.microsoft.com/office/word/2010/wordml';
    const { file } = await setup(`<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:pPr><w:rPr><f:ligatures xmlns:f="${ns}" f:val="standard"${unknown ? ' future="retained"' : ''}/></w:rPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>`);
    if (file.content.kind !== 'word') throw Error('Word fixture');
    const original = file.original!.data.slice(0);
    const editor = new Editor({ extensions: wordExtensions(), content: file.content.html });
    try {
      editor.commands.setTextSelection(editor.state.doc.firstChild!.nodeSize + 1);
      expect(editor.commands.joinBackward()).toBe(true);
      file.content.html = editor.getHTML();
      if (unknown) await expect(exported(file)).rejects.toThrow(/document structures/);
      else expect(descendants((await exported(file)).doc, 'p')[0].textContent).toBe('FirstSecond');
      expect(file.original!.data).toEqual(original);
    } finally { editor.destroy(); }
  });
  it.each(['split', 'join'] as const)(
    'exports a live section %s with exact source properties and atomic Undo',
    async (action) => {
      const { file } = await setup(
        '<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rInherited"/><w:pgSz w:w="10000" w:h="14000"/></w:sectPr></w:pPr><w:r><w:t>AlphaBeta</w:t></w:r></w:p><w:p><w:r><w:t>Gamma</w:t></w:r></w:p>',
      );
      if (file.content.kind !== 'word') throw Error();
      const source = file.content.docxStructure;
      const editor = new Editor({
        extensions: [...wordExtensions(), WordEditorSections.configure({ source })],
        content: file.content.html,
      });
      try {
        updateWordSectionSource(editor, source);
        editor.commands.setTextSelection(action === 'split' ? 6 : 12);
        if (action === 'split') editor.commands.splitBlock();
        else editor.commands.joinBackward();
        file.content.html = editor.getHTML();
        file.content.sectionState = editor.state.doc.attrs.wordSectionState;
        const { doc } = await exported(file);
        const sections = descendants(doc, 'sectPr');
        expect(sections).toHaveLength(action === 'split' ? 2 : 1);
        if (action === 'split') {
          expect(sections[0].parentElement!.parentElement!.textContent).toBe('Beta');
          expect(val(descendants(sections[0], 'pgSz')[0], 'w')).toBe('10000');
        } else expect(descendants(doc, 'p')[0].textContent).toBe('AlphaBetaGamma');
        expect(file.content.docxStructure).toBe(source);
        editor.commands.undo();
        file.content.html = editor.getHTML();
        file.content.sectionState = editor.state.doc.attrs.wordSectionState;
        const restored = await exported(file);
        expect(descendants(restored.doc, 'sectPr')).toHaveLength(2);
        expect(
          descendants(restored.doc, 'sectPr')[0].parentElement!.parentElement!.textContent,
        ).toBe('AlphaBeta');
      } finally {
        editor.destroy();
      }
    },
  );
  it('rejects a real editor split at a section end instead of exporting the break before its second fragment', async () => {
    const { file } = await setup(
      '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>AlphaBeta</w:t></w:r></w:p><w:p><w:r><w:t>Gamma</w:t></w:r></w:p>',
    );
    if (file.content.kind !== 'word') throw Error();
    const originalBytes = file.original!.data.slice(0);
    const sourceStructure = JSON.stringify(file.content.docxStructure);
    const editor = new Editor({ extensions: wordExtensions(), content: file.content.html });
    try {
      editor.commands.setTextSelection(6);
      editor.commands.splitBlock();
      file.content.html = editor.getHTML();
      const edited = file.content.html;
      await expect(exportRetainedDocument(file)).rejects.toThrow(
        'splitting a section-ending paragraph',
      );
      expect(file.original!.data).toEqual(originalBytes);
      expect(JSON.stringify(file.content.docxStructure)).toBe(sourceStructure);
      expect(file.content.html).toBe(edited);
      editor.commands.undo();
      file.content.html = editor.getHTML();
      const { doc } = await exported(file);
      expect(descendants(doc, 'sectPr')).toHaveLength(2);
      expect(descendants(doc, 'p')[0].textContent).toBe('AlphaBeta');
    } finally {
      editor.destroy();
    }
  });
  it('edits soft breaks and tabs in existing paragraphs while retaining unaffected run properties and parts', async () => {
    const { file, zip: original } = await setup(
      '<w:p><w:r><w:rPr><w:b/><w:lang w:val="de-DE"/></w:rPr><w:t>Before</w:t><w:br/><w:t>After</w:t><w:tab/><w:t>Tab</w:t></w:r></w:p>',
    );
    const view = html(file),
      p = view.dom.querySelector('p')!;
    expect(p.querySelector('br')).not.toBeNull();
    p.innerHTML = p.innerHTML.replace('Before', 'Before revised').replace('After', 'After revised');
    view.save();
    const { zip, doc, bytes } = await exported(file);
    expect(descendants(doc, 'br')).toHaveLength(1);
    expect(descendants(doc, 'tab')).toHaveLength(1);
    expect(
      descendants(doc, 'rPr').every(
        (pr) => descendants(pr, 'b').length === 1 && val(descendants(pr, 'lang')[0]) === 'de-DE',
      ),
    ).toBe(true);
    for (const path of Object.keys(original.files))
      if (!original.files[path].dir && path !== 'word/document.xml')
        expect(await zip.file(path)!.async('uint8array'), path).toEqual(
          await original.file(path)!.async('uint8array'),
        );
    expect((await readDocx(bytes)).content.html).toContain('Before revised');
  });
  it('inserts and removes soft breaks without turning them into plain XML text', async () => {
    const { file } = await setup('<w:p><w:r><w:t>Alpha beta</w:t></w:r></w:p>');
    const view = html(file);
    view.dom.querySelector('p')!.innerHTML = 'Alpha<br>beta<br>Gamma';
    view.save();
    const added = await exported(file);
    expect(descendants(added.doc, 'br')).toHaveLength(2);
    file.original!.data = added.bytes;
    file.content = (await readDocx(added.bytes)).content;
    const next = html(file);
    next.dom.querySelector('p')!.innerHTML = 'Alpha beta Gamma';
    next.save();
    expect(descendants((await exported(file)).doc, 'br')).toHaveLength(0);
  });
  it('does not replace clearing breaks with ordinary soft breaks', async () => {
    const { file } = await setup(
      '<w:p><w:r><w:t>Before</w:t><w:br w:clear="all"/><w:t>After</w:t></w:r></w:p>',
    );
    const view = html(file);
    view.dom.querySelector('p')!.innerHTML = 'Changed<br>After';
    view.save();
    await expect(exportRetainedDocument(file)).rejects.toThrow();
  });
  it('retains inline page boundaries and source run properties while editing surrounding text', async () => {
    const { file, zip: original } = await setup(
      '<w:p><w:r><w:rPr><w:b/><w:lang w:val="de-DE"/></w:rPr><w:t>Before</w:t><w:br w:type="page"/><w:t>After</w:t></w:r></w:p>',
    );
    const view = html(file),
      p = view.dom.querySelector('p')!;
    expect(wordJSON(p.outerHTML).content![0].content!.some((n) => n.type === 'wordPageBreak')).toBe(
      true,
    );
    p.innerHTML = p.innerHTML.replace('Before', 'Before revised');
    view.save();
    const { doc, zip, bytes } = await exported(file);
    expect(descendants(doc, 'br').map((e) => val(e, 'type'))).toEqual(['page']);
    expect(
      descendants(doc, 'rPr').every(
        (pr) => descendants(pr, 'b').length === 1 && val(descendants(pr, 'lang')[0]) === 'de-DE',
      ),
    ).toBe(true);
    for (const path of Object.keys(original.files))
      if (!original.files[path].dir && path !== 'word/document.xml')
        expect(await zip.file(path)!.async('uint8array')).toEqual(
          await original.file(path)!.async('uint8array'),
        );
    expect((await readDocx(bytes)).content.html).toContain('data-word-page-break');
  });
  it('inserts and deletes explicit page breaks in retained paragraphs', async () => {
    const { file } = await setup('<w:p><w:r><w:t>Before After</w:t></w:r></w:p>');
    const view = html(file);
    view.dom.querySelector('p')!.innerHTML = 'Before<span data-word-page-break="true"></span>After';
    view.save();
    const added = await exported(file);
    expect(descendants(added.doc, 'br').map((e) => val(e, 'type'))).toEqual(['page']);
    file.original!.data = added.bytes;
    file.content = (await readDocx(added.bytes)).content;
    const next = html(file);
    next.dom.querySelector('[data-word-page-break]')!.remove();
    next.save();
    expect(descendants((await exported(file)).doc, 'br')).toHaveLength(0);
  });
  it('writes page boundaries for new DOCX documents without converting them to line breaks', async () => {
    const file = newFile('word', 'Page boundaries', {
      kind: 'word',
      paper: 'a4',
      margin: 'normal',
      html: '<p>First<span data-word-page-break="true"></span>Second</p>',
    });
    const blob = await exportOffice(file);
    const bytes = await new Promise<ArrayBuffer>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.readAsArrayBuffer(blob);
    });
    const zip = await JSZip.loadAsync(bytes);
    expect(
      descendants(wordXml(await zip.file('word/document.xml')!.async('string')), 'br').map((e) =>
        val(e, 'type'),
      ),
    ).toEqual(['page']);
    expect((await readDocx(bytes)).content.html).toContain('data-word-page-break');
  });
  it('imports inherited paragraph controls and rejects an unimplemented hanging-style conversion explicitly', async () => {
    const { file, zip } = await setup(
      '<w:p><w:pPr><w:pStyle w:val="Hanging"/></w:pPr><w:r><w:t>Inherited layout</w:t></w:r></w:p>',
    );
    zip.file(
      'word/styles.xml',
      `<w:styles xmlns:w="${WORD_NS}"><w:style w:type="paragraph" w:styleId="Hanging"><w:name w:val="Hanging"/><w:pPr><w:keepNext/><w:keepLines/><w:ind w:left="720" w:hanging="360"/></w:pPr></w:style></w:styles>`,
    );
    const data = await zip.generateAsync({ type: 'arraybuffer' });
    file.original!.data = data;
    file.content = (await readDocx(data)).content;
    expect(wordJSON(file.content.html).content![0].attrs).toMatchObject({
      indentStart: '36pt',
      firstLineIndent: '-18pt',
      keepNext: true,
      keepLines: true,
    });
    const view = html(file);
    view.dom.querySelector('p')!.style.textIndent = '18pt';
    view.save();
    await expect(exportRetainedDocument(file)).rejects.toThrow('inherited hanging indent');
  });
  it('writes paragraph indents and pagination controls for newly created DOCX files', async () => {
    const file = newFile('word', 'Paragraph settings', {
      kind: 'word',
      paper: 'a4',
      margin: 'normal',
      html: '<p style="margin-inline-start:36pt;margin-inline-end:12pt;text-indent:-18pt;break-after:avoid;break-inside:avoid;break-before:page;orphans:1;widows:1">New paragraph</p>',
    });
    const blob = await exportOffice(file);
    const data = await new Promise<ArrayBuffer>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.readAsArrayBuffer(blob);
    });
    const content = (await readDocx(data)).content;
    expect(wordJSON(content.html).content![0].attrs).toMatchObject({
      indentStart: '36pt',
      indentEnd: '12pt',
      firstLineIndent: '-18pt',
      keepNext: true,
      keepLines: true,
      pageBreakBefore: true,
      widowControl: false,
    });
  });
  it('imports, edits and clears paragraph indents and pagination flags while retaining other parts', async () => {
    const { file, zip: original } = await setup(
      '<w:p><w:pPr><w:keepNext/><w:keepLines/><w:pageBreakBefore w:val="0"/><w:widowControl w:val="0"/><w:ind w:left="720" w:right="240" w:hanging="360"/></w:pPr><w:r><w:t>Paragraph layout</w:t></w:r></w:p>',
    );
    const view = html(file),
      p = view.dom.querySelector('p')!;
    expect(p.style.marginInlineStart).toBe('36pt');
    expect(p.style.textIndent).toBe('-18pt');
    expect(wordJSON(p.outerHTML).content![0].attrs).toMatchObject({
      keepNext: true,
      keepLines: true,
      pageBreakBefore: false,
      widowControl: false,
    });
    p.style.marginInlineStart = '48pt';
    p.style.marginInlineEnd = '24pt';
    p.style.textIndent = '12pt';
    p.style.breakAfter = 'auto';
    p.style.breakInside = 'auto';
    p.style.breakBefore = 'page';
    p.style.orphans = '2';
    p.style.widows = '2';
    view.save();
    const { zip, doc, bytes } = await exported(file);
    const ind = descendants(doc, 'ind')[0];
    expect(val(ind, 'start')).toBe('960');
    expect(val(ind, 'end')).toBe('480');
    expect(val(ind, 'firstLine')).toBe('240');
    expect(val(ind, 'hanging')).toBe('');
    for (const [name, value] of [
      ['keepNext', '0'],
      ['keepLines', '0'],
      ['pageBreakBefore', '1'],
      ['widowControl', '1'],
    ])
      expect(val(descendants(doc, name)[0])).toBe(value);
    for (const path of Object.keys(original.files))
      if (!original.files[path].dir && path !== 'word/document.xml')
        expect(await zip.file(path)!.async('uint8array'), path).toEqual(
          await original.file(path)!.async('uint8array'),
        );
    const reread = await readDocx(bytes);
    expect(wordJSON(reread.content.html).content![0].attrs).toMatchObject({
      indentStart: '48pt',
      indentEnd: '24pt',
      firstLineIndent: '12pt',
      pageBreakBefore: true,
      widowControl: true,
    });
  });
  it('checks retained backup archive expansion limits before parsing source parts', async () => {
    const { file } = await setup();
    const view = new DataView(file.original!.data);
    const end = view.byteLength - 22;
    const directory = view.getUint32(end + 16, true);
    view.setUint32(directory + 24, 101 * 1024 * 1024, true);
    await expect(exportRetainedDocument(file)).rejects.toThrow('100 MB');
  });
  it('maps repeated paragraphs independently and imports authored page size and direction', async () => {
    const { read } = await setup(
      '<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>Same</w:t></w:r></w:p><w:p><w:r><w:t>Same</w:t></w:r></w:p>',
    );
    expect(read.content.paper).toBe('letter');
    expect(read.content.html).toContain('dir="rtl"');
    expect(read.content.html).not.toContain('NOFFICE');
    const nodes = wordJSON(read.content.html).content!;
    expect(nodes.map((n) => n.attrs?.sourceParagraph)).toEqual(['0:0', '0:1']);
  });
  it('changes styled text while retaining source run formatting, headers and every unrelated part', async () => {
    const { file, zip: original } = await setup();
    const edit = html(file);
    edit.dom.querySelector('em')!.textContent = 'new world';
    edit.save();
    const { zip, doc } = await exported(file);
    expect(
      descendants(doc, 't')
        .map((e) => e.textContent)
        .join(''),
    ).toBe('Hello new world');
    expect(descendants(doc, 'lang')[0].getAttributeNS(WORD_NS, 'val')).toBe('de-DE');
    expect(descendants(doc, 'color')[0].getAttributeNS(WORD_NS, 'val')).toBe('185ABD');
    expect(Object.keys(zip.files).sort()).toEqual(Object.keys(original.files).sort());
    for (const [path, part] of Object.entries(original.files))
      if (!part.dir && path !== 'word/document.xml')
        expect(await zip.file(path)!.async('uint8array'), path).toEqual(
          await part.async('uint8array'),
        );
  });
  it('applies a selected formatting change while preserving unmodeled language and source color', async () => {
    const { file } = await setup();
    const edit = html(file);
    edit.dom.querySelector('strong')!.innerHTML = '<u>Hello </u>';
    edit.save();
    const { doc } = await exported(file);
    expect(descendants(doc, 'u')[0].getAttributeNS(WORD_NS, 'val')).toBe('single');
    expect(descendants(doc, 'b')).toHaveLength(1);
    expect(descendants(doc, 'lang')[0].getAttributeNS(WORD_NS, 'val')).toBe('de-DE');
    expect(descendants(doc, 'color')[0].getAttributeNS(WORD_NS, 'val')).toBe('185ABD');
  });
  it('retains bookmark boundaries during a text edit', async () => {
    const { file } = await setup(
      '<w:p><w:bookmarkStart w:id="7" w:name="Here"/><w:r><w:t>Original</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>',
    );
    const edit = html(file);
    edit.dom.querySelector('p span')!.textContent = 'Changed';
    edit.save();
    const { doc } = await exported(file);
    expect(descendants(doc, 't')[0].textContent).toBe('Changed');
    expect(descendants(doc, 'bookmarkStart')[0].getAttributeNS(WORD_NS, 'id')).toBe('7');
    expect(descendants(doc, 'bookmarkEnd')[0].getAttributeNS(WORD_NS, 'id')).toBe('7');
  });
  it('adds rich paragraphs before final section properties without rebuilding the document', async () => {
    const { file } = await setup();
    const edit = html(file);
    edit.dom.body.insertAdjacentHTML(
      'beforeend',
      '<p style="text-align:center"><strong>A &amp; B</strong><br>Next</p>',
    );
    edit.save();
    const { doc } = await exported(file);
    const body = descendants(doc, 'body')[0];
    expect(body.lastElementChild!.localName).toBe('sectPr');
    expect(descendants(doc, 'headerReference')).toHaveLength(1);
    expect(descendants(doc, 'br')).toHaveLength(1);
    expect(descendants(doc, 'jc').at(-1)!.getAttributeNS(WORD_NS, 'val')).toBe('center');
    expect(
      descendants(doc, 't')
        .map((e) => e.textContent)
        .join(''),
    ).toContain('A & BNext');
  });
  it('refuses editing a field result and retains the original instead of silently flattening it', async () => {
    const { file } = await setup(
      '<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>Today</w:t></w:r></w:fldSimple></w:p>',
    );
    const edit = html(file);
    edit.dom.querySelector('p')!.textContent = 'Tomorrow';
    edit.save();
    await expect(exportRetainedDocument(file)).rejects.toThrow('field');
  });
  it('keeps deliberately added blank paragraphs', async () => {
    const { file } = await setup();
    const edit = html(file);
    edit.dom.body.insertAdjacentHTML('beforeend', '<p></p><p></p>');
    edit.save();
    const { doc } = await exported(file);
    expect(descendants(doc, 'p')).toHaveLength(3);
    expect(descendants(doc, 'body')[0].lastElementChild!.localName).toBe('sectPr');
  });
  it('rotates custom source paper dimensions without replacing them with an A4 preset', async () => {
    const initial = await setup();
    const xml = await initial.zip.file('word/document.xml')!.async('string');
    initial.zip.file(
      'word/document.xml',
      xml.replace('w:w="12240" w:h="15840"', 'w:w="14000" w:h="20000"'),
    );
    const data = await initial.zip.generateAsync({ type: 'arraybuffer' });
    const imported = await readDocx(data);
    const file = newFile('word', 'Custom paper', { ...imported.content, orientation: 'landscape' });
    file.original = { name: 'custom.docx', data };
    const { doc } = await exported(file);
    const size = descendants(doc, 'pgSz')[0];
    expect(size.getAttributeNS(WORD_NS, 'w')).toBe('20000');
    expect(size.getAttributeNS(WORD_NS, 'h')).toBe('14000');
    expect(descendants(doc, 'headerReference')).toHaveLength(1);
  });
  it('does not copy source section breaks when adding a paragraph', async () => {
    const { file } = await setup(
      '<w:p><w:pPr><w:sectPr><w:cols w:num="2"/></w:sectPr></w:pPr><w:r><w:t>Section ending</w:t></w:r></w:p>',
    );
    const edit = html(file);
    edit.dom.body.insertAdjacentHTML('beforeend', '<p>New section text</p>');
    edit.save();
    const { doc } = await exported(file);
    expect(descendants(doc, 'sectPr')).toHaveLength(2);
    expect(descendants(doc, 'cols')).toHaveLength(1);
    expect(descendants(doc, 'p').at(-1)!.textContent).toBe('New section text');
  });
  it('refuses table restructuring while allowing edits within a mapped table paragraph', async () => {
    const { file } = await setup(
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    );
    const edit = html(file);
    edit.dom.querySelector('td p')!.textContent = 'Edited cell';
    edit.save();
    const { doc } = await exported(file);
    expect(descendants(doc, 't')[0].textContent).toBe('Edited cell');
    edit.dom.querySelector('table')!.remove();
    edit.save();
    await expect(exportRetainedDocument(file)).rejects.toThrow('structures');
  });
  it('reads current final settings and changes only live sections, preserving tracked history', async () => {
    const { file, zip } = await setup();
    const doc = wordXml(await zip.file('word/document.xml')!.async('string'));
    const section = descendants(doc, 'sectPr')[0];
    const history = doc.createElementNS(WORD_NS, 'w:sectPrChange');
    history.innerHTML = `<w:sectPr xmlns:w="${WORD_NS}"><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>`;
    section.append(history);
    const oldHistory = new XMLSerializer().serializeToString(history);
    zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
    const data = await zip.generateAsync({ type: 'arraybuffer' });
    file.original!.data = data;
    file.content = (await readDocx(data)).content;
    expect(file.content.orientation).toBe('portrait');
    expect(file.content.paper).toBe('letter');
    expect(file.content.docxStructure!.sections).toHaveLength(1);
    file.content.orientation = 'landscape';
    const exportedDoc = (await exported(file)).doc;
    expect(new XMLSerializer().serializeToString(descendants(exportedDoc, 'sectPrChange')[0])).toBe(
      oldHistory,
    );
    expect(
      descendants(exportedDoc, 'sectPr')[0]
        .getElementsByTagNameNS(WORD_NS, 'pgSz')[0]
        .getAttributeNS(WORD_NS, 'orient'),
    ).toBe('landscape');
  });
  it('rejects source identity tampering and still exports legacy snapshots without metadata', async () => {
    const { file } = await setup();
    if (file.content.kind !== 'word') throw Error();
    file.content.docxStructure!.sections[0].endingParagraph = '0:999';
    await expect(exportRetainedDocument(file)).rejects.toThrow('source-section identities');
    delete file.content.docxStructure;
    const edit = html(file);
    edit.dom.querySelector('p')!.textContent = 'Legacy edit';
    edit.save();
    expect(
      descendants((await exported(file)).doc, 't')
        .map((e) => e.textContent)
        .join(''),
    ).toContain('Legacy edit');
  });
  it('exports explicit page presets even when they equal the approximate imported controls', async () => {
    const { file, zip } = await setup();
    const xml = (await zip.file('word/document.xml')!.async('string')).replace(
      'w:w="12240" w:h="15840"',
      'w:w="10000" w:h="14000"',
    );
    zip.file('word/document.xml', xml);
    file.original!.data = await zip.generateAsync({ type: 'arraybuffer' });
    file.content = (await readDocx(file.original!.data)).content;
    expect(file.content).toMatchObject({ paper: 'a4', margin: 'normal' });
    file.content.pageOverrides = { paper: true, margin: true };
    const { doc } = await exported(file);
    expect(val(descendants(doc, 'pgSz')[0], 'w')).toBe('11906');
    expect(val(descendants(doc, 'pgSz')[0], 'h')).toBe('16838');
    expect(
      ['left', 'right', 'top', 'bottom'].map((side) => val(descendants(doc, 'pgMar')[0], side)),
    ).toEqual(['1440', '1440', '1440', '1440']);
  });
});
