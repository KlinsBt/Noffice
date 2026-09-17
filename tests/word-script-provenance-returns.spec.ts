import {test,expect} from '@playwright/test';
import fs from 'node:fs/promises';
import {createHash} from 'node:crypto';
import reference from './fixtures/native-word-script-provenance.json' with {type:'json'};
import {wordStoryBuildHash} from './word-story-artifacts';

for(const row of reference.rows)test(`Word script provenance ${row.name}: native edit returns`,async({page})=>{
  test.skip(process.env.NOFFICE_SCRIPT_PROVENANCE_RETURN!=='1','Requires independently edited native output');
  const sourceRun=process.env.NOFFICE_SCRIPT_PROVENANCE_SOURCE_RUN||'browser-v5',run=process.env.NOFFICE_SCRIPT_PROVENANCE_RETURN_RUN||'browser-v1';
  if(!/^browser-v\d+$/.test(sourceRun)||!/^browser-v\d+$/.test(run))throw Error('Invalid evidence folder');
  const nativeRoot=`.local/word-script-provenance/${sourceRun}/${row.name}/native-v1`,root=nativeRoot+'/'+run;
  await fs.mkdir(root,{recursive:true});const hash=(bytes:Buffer)=>createHash('sha256').update(bytes).digest('hex');
  const receiptBytes=await fs.readFile(nativeRoot+'/native-report.json'),receipt=JSON.parse(receiptBytes.toString().replace(/^\uFEFF/,''));
  const expected=receipt.stages.find((stage:{stage:string})=>stage.stage==='returned');
  const bytes=await fs.readFile(nativeRoot+'/returned.docx');expect(hash(bytes)).toBe(expected.docxHash);
  expect(receipt.executableHash).toBe(reference.executableHash);
  const name='Returned provenance '+row.name,kind=row.kind==='Headers'?'header':'footer';
  const errors:string[]=[];page.on('pageerror',error=>errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({name:name+'.docx',buffer:bytes,mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const open=async()=>{
    await page.getByRole('button',{name:'Insert',exact:true}).click();await page.getByRole('button',{name:'Headers and footers',exact:true}).click();
    await page.getByRole('button',{name:new RegExp(`^Section ${row.section} — Default ${kind}`)}).click();
  };
  const editor=page.getByRole('textbox',{name:'Header or footer text',exact:true});
  const snapshot=()=>editor.evaluate(element=>{
    type Node={type:string;text?:string;attrs:Record<string,unknown>;content?:Node[];marks?:{type:string;attrs:Record<string,unknown>}[]};
    const p=(element as HTMLElement&{editor:{getJSON():Node}}).editor.getJSON().content![0];
    const read=(text:string,marks:Node['marks']=[],paragraph=false)=>{
      const attrs=marks.find(m=>m.type==='textStyle')?.attrs;
      return {text,family:String(attrs?.fontFamily||p.attrs.paragraphFontFamily).replace(/^["']|["']$/g,''),size:parseFloat(String(attrs?.fontSize||p.attrs.paragraphFontSize)),
        superscript:(paragraph?p.attrs.paragraphScript==='superscript':marks.some(m=>m.type==='superscript'))?-1:0,
        subscript:(paragraph?p.attrs.paragraphScript==='subscript':marks.some(m=>m.type==='subscript'))?-1:0};
    };
    return [...(p.content||[]).flatMap(n=>[...(n.text||'')].map(c=>read(c,n.marks))),read('\r',[],true)];
  });
  await open();await expect(editor).toHaveText('Returned');expect(await snapshot()).toEqual(expected.characters);
  await editor.focus();await page.keyboard.press('Control+End');await page.keyboard.insertText('Z');await expect(editor).toHaveText('ReturnedZ');
  await page.keyboard.press('Control+z');expect(await snapshot()).toEqual(expected.characters);
  await page.keyboard.press('Control+y');await expect(editor).toHaveText('ReturnedZ');await page.keyboard.press('Control+z');
  expect(await snapshot()).toEqual(expected.characters);await page.getByRole('button',{name:'Cancel',exact:true}).click();
  const outputs:Record<string,string>={};
  const download=async(file:string,label:string)=>{
    await page.getByRole('button',{name:'Export',exact:true}).click();const pending=page.waitForEvent('download');
    await page.getByRole('button',{name:label,exact:true}).click();await(await pending).saveAs(root+'/'+file);outputs[file]=hash(await fs.readFile(root+'/'+file));
  };
  await download('returned.docx','DOCX file Editable in Microsoft Word');expect(await fs.readFile(root+'/returned.docx')).toEqual(bytes);
  await download('returned.pdf','PDF file');await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();
  await page.getByRole('button',{name,exact:true}).click();await download('reloaded.pdf','PDF file');
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles(root+'/returned.docx');await open();expect(await snapshot()).toEqual(expected.characters);
  expect(errors).toEqual([]);await fs.writeFile(root+'/browser-report.json',JSON.stringify({outputs,errors,sourceHash:hash(bytes),nativeReceiptHash:hash(receiptBytes),
    buildHash:await wordStoryBuildHash(),referenceHash:hash(await fs.readFile('tests/fixtures/native-word-script-provenance.json')),
    testHash:hash(await fs.readFile('tests/word-script-provenance-returns.spec.ts'))},null,2));
});
