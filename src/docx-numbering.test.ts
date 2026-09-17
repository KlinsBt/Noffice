import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { Editor } from '@tiptap/core';
import { docxNumbering } from './docx-numbering';
import { wordXml, WORD_NS } from './word-xml';
import { readDocx } from './docx-import';
import { contentSchema, newFile } from './model';
import { contentFingerprint } from './office-preservation';
import { hydrateWordStructure, needsWordStructure } from './word-structure';
import { wordExtensions } from './word-extensions';
import { WordEditorSections, updateWordSectionSource, wordSectionMap } from './word-editor-sections';
import { resolveWordLists } from './word-list-resolve';
import { resolveWordSections } from './word-section-layout';
import { sectionSurfaceInput } from './word-section-surfaces';

vi.mock('mammoth', async original => {
  const actual = await original<typeof import('mammoth')>();
  return { ...actual, convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
    actual.convertToHtml({ buffer: Buffer.from(input.arrayBuffer) }, options as Parameters<typeof actual.convertToHtml>[1]) };
});
const xml = `<w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num></w:numbering>`;
const properties = (extra = '') => wordXml(`<w:pPr xmlns:w="${WORD_NS}"><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr>${extra}</w:pPr>`).documentElement;
it('resolves instance identity, start, marker font and twip geometry without mutating properties', () => {
  const p = properties(), before = new XMLSerializer().serializeToString(p);
  expect(docxNumbering(xml)(p)).toEqual({status:'resolved',definition:{numId:'7',level:0,start:3,
    format:'decimal',text:'%1.',left:720,hanging:360,font:'Arial',size:12}});
  expect(new XMLSerializer().serializeToString(p)).toBe(before);
  const baseline=properties();baseline.getElementsByTagNameNS(WORD_NS,'rPr')[0].appendChild(
    wordXml(`<w:vertAlign xmlns:w="${WORD_NS}" w:val="baseline"/>`).documentElement);
  expect(docxNumbering(xml)(baseline)?.status).toBe('resolved');
  expect(docxNumbering(xml.replace('</w:num>', '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="9"/></w:lvlOverride></w:num>'))(p))
    .toMatchObject({status:'resolved',definition:{start:9}});
});
it('uses a complete level override and marker-specific font and size', () => {
  const override='<w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\u2022"/><w:pPr><w:ind w:left="960" w:hanging="480"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:lvl></w:lvlOverride>';
  expect(docxNumbering(xml.replace('</w:num>',override+'</w:num>'))(properties()))
    .toMatchObject({status:'resolved',definition:{start:1,format:'bullet',text:'\u2022',left:960,hanging:480,font:'Calibri',size:11}});
});
for (const [name, replace, withValue] of [
  ['dangling instance', 'w:numId="7"','w:numId="8"'],
  ['dangling abstract','<w:abstractNumId w:val="4"/>','<w:abstractNumId w:val="5"/>'],
  ['duplicate instance','</w:numbering>','<w:num w:numId="7"/></w:numbering>'],
  ['duplicate abstract','</w:numbering>','<w:abstractNum w:abstractNumId="4"/></w:numbering>'],
  ['duplicate format','<w:numFmt w:val="decimal"/>','<w:numFmt w:val="decimal"/><w:numFmt w:val="decimal"/>'],
  ['custom number format','w:val="decimal"','w:val="upperRoman"'],
  ['unqualified bullet','w:val="%1."','w:val="%1)"'],
  ['negative indent','w:left="720"','w:left="-1"'],
  ['overflow indent','w:left="720"','w:left="999999999999999999999"'],
  ['hanging outside body','w:hanging="360"','w:hanging="800"'],
  ['character indent','w:left="720"','w:leftChars="720"'],
  ['picture marker','</w:lvl>','<w:lvlPicBulletId w:val="1"/></w:lvl>'],
  ['linked style','</w:abstractNum>','<w:numStyleLink w:val="Other"/></w:abstractNum>'],
  ['right alignment','w:val="left"','w:val="right"'],
  ['space suffix','</w:lvl>','<w:suff w:val="space"/></w:lvl>'],
  ['bold marker','</w:lvl>','<w:rPr><w:b/></w:rPr></w:lvl>'],
  ['hidden level text','<w:lvlText w:val="%1."/>','<w:lvlText w:val="%1." w:null="1"/>'],
  ['custom format attribute','<w:numFmt w:val="decimal"/>','<w:numFmt w:val="decimal" w:format="custom"/>'],
] as const) it(`retains an explicit unsupported result for ${name}`, () => {
  expect(docxNumbering(xml.replace(replace,withValue))(properties())?.status).toBe('unsupported');
});
it('rejects paragraph indent overrides, absent parts, nested levels and invalid backup metadata', () => {
  expect(docxNumbering(xml)(properties('<w:ind w:left="900"/>'))?.status).toBe('unsupported');
  expect(docxNumbering()(properties())?.status).toBe('unsupported');
  const bold=properties();bold.getElementsByTagNameNS(WORD_NS,'rPr')[0].appendChild(bold.ownerDocument.createElementNS(WORD_NS,'w:b'));
  expect(docxNumbering(xml)(bold)?.status).toBe('unsupported');
  const p=properties();p.getElementsByTagNameNS(WORD_NS,'ilvl')[0].setAttributeNS(WORD_NS,'w:val','1');
  expect(docxNumbering(xml)(p)?.status).toBe('unsupported');
  const content={kind:'word',html:'<p/>',paper:'letter',margin:'normal',numbering:{version:1,paragraphs:[
    {source:'0:1',numbering:docxNumbering(xml)(properties())},
    {source:'0:1',numbering:docxNumbering(xml)(properties())}]}};
  expect(contentSchema.safeParse(content).success).toBe(false);
});
for (const kind of ['numbered','bulleted']) it(`migrates ${kind} source metadata without replacing edited HTML or original identity`, async () => {
  const data=Uint8Array.from(readFileSync(`tests/fixtures/word-section-containers/${kind}.docx`)).buffer;
  const {content}=await readDocx(data);
  expect(content.numbering?.paragraphs).toHaveLength(3);
  expect(content.numbering?.paragraphs.every(p=>p.numbering.status==='resolved')).toBe(true);
  const legacy={...content,numbering:undefined}, file=newFile('word','legacy',legacy);
  file.original={name:'legacy.docx',data,contentFingerprint:await contentFingerprint(legacy)};
  expect(await contentFingerprint(content)).toBe(file.original.contentFingerprint);
  file.content={...legacy,html:legacy.html.replace('Beta item.','Edited Beta item.')};
  expect(needsWordStructure(file)).toBe(true);
  const migrated=await hydrateWordStructure(file);
  expect(migrated.content).toEqual({...file.content,numbering:content.numbering});
  expect(migrated.original).toBe(file.original);expect(migrated.revision).toBe(file.revision);
  expect(needsWordStructure(migrated)).toBe(false);
  const editor=new Editor({extensions:[...wordExtensions({numbering:()=>content.numbering}),WordEditorSections],content:content.html});
  try {
    updateWordSectionSource(editor,content.docxStructure);
    const resolved=resolveWordLists(editor.state.doc,content.numbering)!;
    expect(resolved,JSON.stringify(editor.getJSON())).not.toBeNull();
    expect([...resolved.values()].map(p=>p.marker)).toEqual(kind==='numbered'?['1.','2.','3.']:['\u2022','\u2022','\u2022']);
    const input=sectionSurfaceInput(editor.state.doc,wordSectionMap(editor.state),resolveWordSections(content),undefined,undefined,content.numbering)!;
    expect(input.blocks).toHaveLength(5);
    expect(input.blocks.filter(p=>p.list).map(p=>[p.indentStart,p.width])).toEqual([[48,576],[48,576],[48,576]]);
    expect(editor.getHTML()).not.toContain('data-word-list-marker');
    expect(editor.getHTML()).not.toContain('word-source-list');
  } finally {editor.destroy();}
});
