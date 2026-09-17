import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { readDocx, WORD_NS } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { newFile, contentSchema } from './model';
import { contentFingerprint } from './office-preservation';
import { wordExtensions } from './word-extensions';
import { WordEditorSections, updateWordSectionSource, wordSectionMap } from './word-editor-sections';
import { insertWordSectionBreak } from './word-section-insert';
import { sectionParagraphs, type WordSectionState } from './word-section-breaks';
import { wordSectionStartSchema } from './word-section-identity';
import { resolveWordSections } from './word-section-layout';
import { exportOffice } from './formats';
import reference from '../tests/fixtures/word-section-containers/reference.json';
import freshReference from '../tests/fixtures/word-section-containers/fresh-reference.json';
import { resolveWordLists } from './word-list-resolve';
import { sectionSurfaceInput } from './word-section-surfaces';

vi.mock('mammoth', async original => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});
const bytes = (blob: Blob) => new Promise<ArrayBuffer>((resolve, reject) => {
  const reader = new FileReader(); reader.onload = () => resolve(reader.result as ArrayBuffer);
  reader.onerror = () => reject(reader.error); reader.readAsArrayBuffer(blob);
});
const extensions = () => [...wordExtensions(), WordEditorSections];
const nativeText = (editor: Editor) => {
  const state = editor.state.doc.attrs.wordSectionState as WordSectionState;
  const endings = new Set(state.breaks.map(b => b.paragraph));
  return sectionParagraphs(editor.state.doc).map((p,i) => p.node.textContent + (endings.has(i) ? '\f' : '\r')).join('');
};
for (const sample of [...reference.rows.filter(r => !r.name.startsWith('table')), ...freshReference.rows])
  it(`preserves native list continuity, selection, history and both writers: ${sample.name}`, async () => {
    const data = Uint8Array.from(readFileSync(sample.source)).buffer;
    expect(createHash('sha256').update(new Uint8Array(data)).digest('hex')).toBe(sample.sourceHash);
    const { content } = await readDocx(data), file = newFile('word',sample.name,content);
    file.original = { name: sample.name+'.docx', data, contentFingerprint: await contentFingerprint(content) };
    const sourceJson = JSON.stringify(content);
    const editor = new Editor({extensions:extensions(),content:content.html,parseOptions:{preserveWhitespace:true}});
    try {
      updateWordSectionSource(editor,content.docxStructure);
      const paragraphs = sectionParagraphs(editor.state.doc), selection=sample.authoredSelection;
      editor.commands.setTextSelection({from:paragraphs[selection.fromParagraph-1].from+1+selection.fromOffset,
        to:paragraphs[selection.toParagraph-1].from+1+selection.toOffset});
      const before = editor.getJSON(), from = editor.state.selection.from;
      expect(insertWordSectionBreak(editor,content,wordSectionStartSchema.parse(sample.kind))).toBe(true);
      expect(editor.state.selection.from).toBe(from+4);
      expect(editor.state.selection.empty).toBe(true);
      expect(nativeText(editor)).toBe(sample.after);
      const lists=resolveWordLists(editor.state.doc,content.numbering)!;
      expect(lists).not.toBeNull();
      for (const [ordinal,p] of sectionParagraphs(editor.state.doc).entries()) {
        const list=lists.get(p.from), native=sample.containers.paragraphs[ordinal];
        if (!list) continue;
        expect(list.definition.left/20).toBe(native.left);
        expect(-list.definition.hanging/20).toBe(native.first);
        if (p.node.content.size) expect(list.marker).toBe(native.listString);
      }
      expect(sectionSurfaceInput(editor.state.doc,wordSectionMap(editor.state),
        resolveWordSections({...content,sectionState:editor.state.doc.attrs.wordSectionState}),
        undefined,undefined,content.numbering)?.blocks).toHaveLength(sample.containers.paragraphs.length);
      expect(editor.state.doc.childCount).toBe(before.content!.length); // section splitting retains wrappers
      const after = editor.getJSON();
      editor.commands.undo(); expect(editor.getJSON()).toEqual(before);
      editor.commands.redo(); expect(editor.getJSON()).toEqual(after);
      const saved = {...content, html:editor.getHTML(), sectionState:editor.state.doc.attrs.wordSectionState as WordSectionState};
      expect(contentSchema.parse(saved)).toEqual(saved);
      const reload = new Editor({extensions:extensions(),content:saved.html,parseOptions:{preserveWhitespace:true}});
      try {
        updateWordSectionSource(reload,content.docxStructure,saved.sectionState);
        expect(wordSectionMap(reload.state)).toEqual(wordSectionMap(editor.state));
        expect(nativeText(reload)).toBe(sample.after);
      } finally { reload.destroy(); }
      file.content = saved;
      const output=await bytes(await exportRetainedDocument(file)), imported=await readDocx(output);
      expect(imported.content.numbering?.paragraphs.every(p=>p.numbering.status==='resolved')).toBe(true);
      const a=await JSZip.loadAsync(data), b=await JSZip.loadAsync(output);
      for (const path of Object.keys(a.files).filter(p=>!a.files[p].dir && p !== 'word/document.xml' && p !== 'word/settings.xml'))
        expect(await b.file(path)!.async('uint8array'),path).toEqual(await a.file(path)!.async('uint8array'));
      expect(resolveWordSections(imported.content).map(s=>s.start)).toEqual(['nextPage',sample.kind]);
      const returned=new Editor({extensions:extensions(),content:imported.content.html,parseOptions:{preserveWhitespace:true}});
      try {updateWordSectionSource(returned,imported.content.docxStructure);expect(nativeText(returned)).toBe(sample.after);}
      finally{returned.destroy();}
      const fresh = newFile('word','Fresh list',{...saved,stories:undefined});
      const freshOutput=await bytes(await exportOffice(fresh)), freshZip=await JSZip.loadAsync(freshOutput);
      const xml = new DOMParser().parseFromString(await freshZip.file('word/document.xml')!.async('string'),'application/xml');
      const listParagraphs=[...xml.getElementsByTagNameNS(WORD_NS,'numPr')];
      expect(listParagraphs).toHaveLength(lists.size);
      const nums=listParagraphs.map(p=>p.getElementsByTagNameNS(WORD_NS,'numId')[0].getAttributeNS(WORD_NS,'val'));
      expect(new Set(nums).size).toBe(new Set([...lists.values()].map(p=>p.definition.numId)).size);
      expect(xml.getElementsByTagNameNS(WORD_NS,'sectPr')).toHaveLength(2);
      const freshImported = (await readDocx(freshOutput)).content;
      expect(freshImported.numbering?.paragraphs).toHaveLength(lists.size);
      for (const [index,entry] of freshImported.numbering!.paragraphs.entries()) {
        expect(entry.numbering.status,JSON.stringify(entry.numbering)).toBe('resolved');
        const {numId:_,...expected}=[...lists.values()][index].definition;
        if (entry.numbering.status === 'resolved') expect(entry.numbering.definition).toMatchObject(expected);
      }
      expect(JSON.stringify(content)).toBe(sourceJson);
      expect(new Uint8Array(file.original.data)).toEqual(new Uint8Array(data));
    } finally {editor.destroy();}
  });

it('rejects nested and multi-paragraph list items and table selections atomically', () => {
  for (const html of ['<ol><li><p>First</p><p>More</p></li></ol>',
    '<ol><li><p>First</p><ol><li><p>Nested</p></li></ol></li></ol>',
    '<table><tr><td><p>Cell</p></td></tr></table>']) {
    const content={kind:'word' as const,html,paper:'letter' as const,margin:'normal' as const};
    const editor=new Editor({extensions:extensions(),content:html});
    try {
      editor.commands.setTextSelection(sectionParagraphs(editor.state.doc)[0].from+2);
      const before=editor.getJSON(), undo=editor.can().undo();
      expect(insertWordSectionBreak(editor,content,'nextPage')).toBe(false);
      expect(editor.getJSON()).toEqual(before);expect(editor.can().undo()).toBe(undo);
    } finally {editor.destroy();}
  }
});

it('recomputes empty ending markers after typing and history without saving view-only numbering', () => {
  const content={kind:'word' as const,html:'<p>Before.</p><ol><li><p>Alpha</p></li><li><p>Beta</p></li><li><p>Gamma</p></li></ol><p>After.</p>',paper:'letter' as const,margin:'normal' as const};
  const editor=new Editor({extensions:extensions(),content:content.html});
  try {
    editor.commands.setTextSelection(sectionParagraphs(editor.state.doc)[2].from+1);
    expect(insertWordSectionBreak(editor,content,'continuous')).toBe(true);
    const snapshot=editor.getJSON();
    const items=()=>[...editor.view.dom.querySelectorAll('li')].map(li=>({value:li.getAttribute('value'),hidden:li.hasAttribute('data-word-section-list-empty')}));
    expect(items()).toEqual([{value:'1',hidden:false},{value:null,hidden:true},{value:'2',hidden:false},{value:'3',hidden:false}]);
    expect(editor.getHTML()).not.toMatch(/data-word-section-list-empty|<li value=/);
    editor.commands.setTextSelection(sectionParagraphs(editor.state.doc)[2].from+1);
    editor.commands.insertContent('X');
    expect(items().every(item=>!item.hidden)).toBe(true);
    editor.commands.undo();expect(editor.getJSON()).toEqual(snapshot);
    expect(items()[3].value).toBe('3');
    editor.commands.redo();expect(items().every(item=>!item.hidden)).toBe(true);
  } finally {editor.destroy();}
});

it('exports repeated fresh section breaks within one list without losing enclosing numbering', async () => {
  const content={kind:'word' as const,html:'<p>Before.</p><ol><li><p>Alpha</p></li><li><p>Beta</p></li><li><p>Gamma</p></li></ol><p>After.</p>',paper:'letter' as const,margin:'normal' as const};
  const editor=new Editor({extensions:extensions(),content:content.html});
  try {
    editor.commands.setTextSelection(sectionParagraphs(editor.state.doc)[1].from+3);
    expect(insertWordSectionBreak(editor,content,'nextPage')).toBe(true);
    editor.commands.setTextSelection(sectionParagraphs(editor.state.doc)[3].from+3);
    expect(insertWordSectionBreak(editor,content,'oddPage')).toBe(true);
    const saved={...content,html:editor.getHTML(),sectionState:editor.state.doc.attrs.wordSectionState as WordSectionState};
    const output=await bytes(await exportOffice(newFile('word','Repeated list sections',saved)));
    const zip=await JSZip.loadAsync(output);
    const xml=new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('string'),'application/xml');
    expect(xml.getElementsByTagNameNS(WORD_NS,'numPr')).toHaveLength(5);
    expect(xml.getElementsByTagNameNS(WORD_NS,'sectPr')).toHaveLength(3);
    const imported=await readDocx(output);
    const loaded=new Editor({extensions:extensions(),content:imported.content.html});
    try {updateWordSectionSource(loaded,imported.content.docxStructure);expect(nativeText(loaded)).toBe(nativeText(editor));}
    finally{loaded.destroy();}
  } finally {editor.destroy();}
});
