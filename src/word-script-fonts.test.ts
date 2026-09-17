import {expect,it,vi} from 'vitest';
import {Editor} from '@tiptap/core';
import {readFileSync} from 'node:fs';
import {createHash} from 'node:crypto';
import {readDocx} from './docx-import';
import {wordExtensions} from './word-extensions';
import native from '../tests/fixtures/native-word-script-fonts.json';

vi.mock('mammoth',async(original)=>{
  const actual=await original<typeof import('mammoth')>();
  return {...actual,convertToHtml:(input:{arrayBuffer:ArrayBuffer},options:unknown)=>
    actual.convertToHtml({buffer:Buffer.from(input.arrayBuffer)},options as Parameters<typeof actual.convertToHtml>[1])};
});

for(const row of native.rows)it(`keeps native nominal fonts and script state for ${row.name}`,async()=>{
  const bytes=readFileSync('tests/fixtures/'+row.sourceFile);expect(createHash('sha256').update(bytes).digest('hex')).toBe(row.stages[0].docxHash);
  const {content}=await readDocx(new Uint8Array(bytes).buffer);
  const html=row.kind==='Body'?content.html:content.stories!.parts.find(p=>p.kind===(row.kind==='Headers'?'header':'footer')&&
    new DOMParser().parseFromString(p.html,'text/html').querySelector('p')?.textContent==='A\tB')!.html;
  const editor=new Editor({extensions:wordExtensions(),content:html});
  const snapshot=()=>{
    const p=editor.getJSON().content![0];
    const read=(text:string,marks:typeof p.marks=[])=>{
      const attrs=marks?.find(m=>m.type==='textStyle')?.attrs;
      return {text,family:String(attrs?.fontFamily||p.attrs!.paragraphFontFamily).replace(/^["']|["']$/g,''),
        size:parseFloat(String(attrs?.fontSize||p.attrs!.paragraphFontSize)),superscript:marks?.some(m=>m.type==='superscript')?-1:0,subscript:marks?.some(m=>m.type==='subscript')?-1:0};
    };
    return [...p.content!.flatMap(n=>[...(n.type==='wordTab'?'\t':'text' in n?n.text:'')].map(c=>read(c,n.marks))),read('\r')];
  };
  try{
    expect(snapshot()).toEqual(row.stages[0].characters);editor.commands.setTextSelection({from:3,to:4});
    const before=editor.getJSON();
    if(row.script==='Superscript')editor.chain().unsetSubscript().toggleSuperscript().run();else editor.chain().unsetSuperscript().toggleSubscript().run();
    expect(snapshot()).toEqual(row.stages[1].characters);const after=editor.getJSON();
    editor.commands.undo();expect(editor.getJSON()).toEqual(before);editor.commands.redo();expect(editor.getJSON()).toEqual(after);
  }finally{editor.destroy();}
});
