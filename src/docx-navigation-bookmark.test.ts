import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { Editor } from '@tiptap/core';
import { collapsedWordNavigationAt } from './docx-navigation-bookmark';
import { WORD_NS, wordXml, descendants, child, val, wElement } from './word-xml';
import { readDocx } from './docx-import';
import { exportRetainedDocument } from './docx-preserve';
import { wordExtensions } from './word-extensions';
import { newFile } from './model';
import { contentFingerprint } from './office-preservation';

vi.mock('mammoth',async original=>{const actual=await original<typeof import('mammoth')>();return {...actual,
  convertToHtml:(input:{arrayBuffer:ArrayBuffer},options:unknown)=>actual.convertToHtml({buffer:Buffer.from(input.arrayBuffer)},options as Parameters<typeof actual.convertToHtml>[1])};});
const bytes=(blob:Blob)=>new Promise<ArrayBuffer>((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(r.result as ArrayBuffer);r.onerror=()=>reject(r.error);r.readAsArrayBuffer(blob);});

for(const malformed of ['named','range','duplicate','column','foreign','field'])
  it(`does not move an unqualified navigation bookmark: ${malformed}`,()=>{
    const name=malformed==='named'?'UserAnchor':'_GoBack';
    const tags=`<w:bookmarkStart w:id="0" w:name="${name}"${malformed==='column'?' w:colFirst="0"':''}${malformed==='foreign'?' x:opaque="1"':''}/>${malformed==='range'?'<w:r><w:t>Q</w:t></w:r>':''}<w:bookmarkEnd w:id="0"/>`;
    const doc=wordXml(`<w:document xmlns:w="${WORD_NS}" xmlns:x="urn:test"><w:body><w:p><w:r>${malformed==='field'?'<w:fldChar w:fldCharType="begin"/>':''}<w:t>X</w:t></w:r>${tags}<w:r><w:t>Beta</w:t></w:r></w:p>${malformed==='duplicate'?'<w:p>'+tags+'</w:p>':''}</w:body></w:document>`);
    expect(collapsedWordNavigationAt(descendants(doc,'p')[0],1)).toEqual([]);
  });

for(const kind of ['numbered','bulleted'])
  it(`exports typed text at an interior collapsed Word navigation anchor in a ${kind} list`,async()=>{
    const zip=await JSZip.loadAsync(readFileSync(`tests/fixtures/word-section-containers/${kind}.docx`));
    const doc=wordXml(await zip.file('word/document.xml')!.async('string'));
    const p=descendants(doc,'p')[2],run=child(p,'r')!;
    p.setAttributeNS(WORD_NS,'w:rsidRDefault','1234ABCD');
    const settings=wordXml(await zip.file('word/settings.xml')!.async('string'));
    const register=wElement(settings,'rsids');register.append(wElement(settings,'rsid',{val:'1234ABCD'}));
    settings.documentElement.append(register);zip.file('word/settings.xml',new XMLSerializer().serializeToString(settings));
    const prefix=run.cloneNode(true) as Element;descendants(prefix,'t')[0].textContent='X';
    p.insertBefore(prefix,run);
    for(const tag of ['bookmarkStart','bookmarkEnd']){
      const marker=doc.createElementNS(WORD_NS,'w:'+tag);marker.setAttributeNS(WORD_NS,'w:id','0');
      if(tag==='bookmarkStart')marker.setAttributeNS(WORD_NS,'w:name','_GoBack');p.insertBefore(marker,run);
    }
    zip.file('word/document.xml',new XMLSerializer().serializeToString(doc));
    const data=await zip.generateAsync({type:'arraybuffer'}), {content}=await readDocx(data);
    const file=newFile('word','Navigation list',content);file.original={name:'Navigation.docx',data,contentFingerprint:await contentFingerprint(content)};
    const editor=new Editor({extensions:wordExtensions({retainedEditRuns:true}),content:content.html});
    try{
      let position=0;editor.state.doc.descendants((n,pos)=>{if(n.attrs.sourceParagraph==='0:2')position=pos+2;});
      editor.commands.setTextSelection(position);editor.commands.insertContent('Y');
      file.content={...content,html:editor.getHTML(),initialEditSession:editor.state.doc.attrs.wordInitialEditSession || undefined};
      const out=await JSZip.loadAsync(await bytes(await exportRetainedDocument(file)));
      const xml=wordXml(await out.file('word/document.xml')!.async('string'));const current=descendants(xml,'p')[2];
      expect(descendants(current,'t').map(t=>t.textContent).join('')).toBe('XYBeta item.');
      const anchor=collapsedWordNavigationAt(current,2);expect(anchor).toHaveLength(2);
      expect(val(anchor[0],'name')).toBe('_GoBack');
      const inserted=anchor[0].previousElementSibling!;expect(descendants(inserted,'t')[0].textContent).toBe('Y');
      expect(val(inserted,'rsidR') || val(current,'rsidRDefault')).toMatch(/^[A-F0-9]{8}$/i);
      expect(await out.file('word/numbering.xml')!.async('string')).toBe(await zip.file('word/numbering.xml')!.async('string'));
      expect(new Uint8Array(file.original.data)).toEqual(new Uint8Array(data));
    }finally{editor.destroy();}
  });
