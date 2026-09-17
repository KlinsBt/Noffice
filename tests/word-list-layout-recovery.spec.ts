import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import type { OfficeFile } from '../src/model';
import { waitWordLayout } from './word-pagination-helpers';
import { wordStoryBuildHash } from './word-story-artifacts';

const hash=(bytes:Buffer)=>createHash('sha256').update(bytes).digest('hex');
for(const kind of ['numbered','bulleted']) for(const mode of ['unchanged','edited','broken-original','font-denied'])
test(`List layout ${kind} ${mode}: recovery preserves saved content and original`,async({page})=>{
  const run=process.env.NOFFICE_LIST_RECOVERY_RUN||'recovery-v1';
  if(!/^recovery-v\d+$/.test(run))throw Error('Invalid recovery run');
  const root=`.local/word-section-containers/${run}/${kind}-${mode}`;await fs.mkdir(root,{recursive:true});
  const source=await fs.readFile(`tests/fixtures/word-section-containers/${kind}.docx`),name=`List recovery ${kind} ${mode}`;
  await page.context().grantPermissions(['local-fonts']);await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({name:name+'.docx',buffer:source,mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const editor=page.getByRole('textbox',{name:'Document text',exact:true});await expect(editor).toBeVisible();
  await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  const before=await page.evaluate(async({name,mode})=>{
    const db=await new Promise<IDBDatabase>((resolve,reject)=>{const r=indexedDB.open('noffice-workspace',1);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
    try{return await new Promise<{revision:number;html:string}>((resolve,reject)=>{
      const tx=db.transaction('files','readwrite'),store=tx.objectStore('files'),r=store.getAll();let result:{revision:number;html:string};
      r.onsuccess=()=>{const file=r.result.find((f:OfficeFile)=>f.name===name) as OfficeFile;if(file.content.kind!=='word')throw Error('Word required');
        if(mode!=='font-denied')delete file.content.numbering;
        if(mode==='edited')file.content.html=file.content.html.replace('Beta item.','Edited Beta item.');
        if(mode==='broken-original')file.original!.data=new Uint8Array(new TextEncoder().encode('preserved invalid original')).buffer;
        result={revision:file.revision,html:file.content.html};store.put(file);};
      tx.oncomplete=()=>resolve(result);tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error);
    });}finally{db.close();}
  },{name,mode});
  await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();await page.getByRole('button',{name,exact:true}).click();
  const capture=async(label:string|RegExp,file:string,recovery=false)=>{
    if(!recovery)await page.getByRole('button',{name:'Export',exact:true}).click();
    const pending=page.waitForEvent('download');await page.getByRole('button',{name:label,exact:typeof label==='string'}).click();await(await pending).saveAs(root+'/'+file);
  };
  if(mode==='broken-original'){
    await expect(page.getByRole('region',{name:'Document recovery'}).getByRole('alert')).toBeVisible();await expect(editor).toHaveCount(0);
    await capture('Export Noffice backup','backup.noffice',true);await capture('Download original DOCX','original.docx',true);
    const backup=JSON.parse(await fs.readFile(root+'/backup.noffice','utf8'));expect(backup.content.html).toBe(before.html);expect(backup.content.numbering).toBeUndefined();
    expect(await fs.readFile(root+'/original.docx','utf8')).toBe('preserved invalid original');
  }else{
    await waitWordLayout(editor);await expect(page.locator('.section-page')).toHaveCount(1);await expect(editor.locator('[data-word-list-marker]')).toHaveCount(3);
    if(mode==='edited')await expect(editor).toContainText('Edited Beta item.');
    if(mode==='font-denied'){
      const model=await editor.evaluate(el=>(el as HTMLElement&{editor:import('@tiptap/core').Editor}).editor.getJSON());
      await page.context().clearPermissions();await page.context().grantPermissions([]);
      await page.getByRole('button',{name:'Export',exact:true}).click();await page.getByRole('button',{name:'PDF file',exact:true}).click();
      await expect(page.getByRole('alert')).toContainText('Font access was not granted');
      expect(await editor.evaluate(el=>(el as HTMLElement&{editor:import('@tiptap/core').Editor}).editor.getJSON())).toEqual(model);
      await page.getByRole('button',{name:'Dismiss error',exact:true}).click();await page.context().grantPermissions(['local-fonts']);
      await capture('PDF file','recovered.pdf');
    }
    await capture(/^Noffice backup/,'backup.noffice');const backup=JSON.parse(await fs.readFile(root+'/backup.noffice','utf8'));
    expect(backup.content.html).toBe(before.html);expect(backup.content.numbering.version).toBe(1);expect(backup.original.base64).toBe(source.toString('base64'));
    await capture('DOCX file Editable in Microsoft Word','export.docx');
    if(mode!=='edited')expect(hash(await fs.readFile(root+'/export.docx'))).toBe(hash(source));
  }
  await fs.writeFile(root+'/report.json',JSON.stringify({kind,mode,sourceHash:hash(source),before,buildHash:await wordStoryBuildHash(),testHash:hash(await fs.readFile('tests/word-list-layout-recovery.spec.ts'))},null,2));
});

for(const scenario of ['unsupported-format','oversized-marker'])test(`List layout ${scenario} retains editable fallback and original bytes`,async({page})=>{
  const zip=await JSZip.loadAsync(await fs.readFile('tests/fixtures/word-section-containers/numbered.docx'));
  let xml=await zip.file('word/numbering.xml')!.async('string');
  xml=scenario==='unsupported-format'?xml.replace('w:val="decimal"','w:val="upperRoman"'):xml.replaceAll('w:val="1"/>','w:val="99999"/>');
  zip.file('word/numbering.xml',xml);const source=await zip.generateAsync({type:'nodebuffer'});
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles({name:scenario+'.docx',buffer:source,mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const editor=page.getByRole('textbox',{name:'Document text',exact:true});await expect(editor).toBeVisible();await waitWordLayout(editor);
  await expect(page.locator('.section-page')).toHaveCount(0);await expect(editor).toContainText('Beta item.');
  await page.getByRole('button',{name:'Export',exact:true}).click();const pending=page.waitForEvent('download');
  await page.getByRole('button',{name:'DOCX file Editable in Microsoft Word',exact:true}).click();const stream=await(await pending).createReadStream(),chunks:Buffer[]=[];
  for await(const chunk of stream!)chunks.push(Buffer.from(chunk));expect(hash(Buffer.concat(chunks))).toBe(hash(source));
});
