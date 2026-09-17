import { test, expect, type Page } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import type { OfficeFile } from '../src/model';
import native from './fixtures/native-word-compatibility.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
const run = process.env.NOFFICE_COMPATIBILITY_RUN || 'browser-v1';
if (!/^browser-v\d+$/.test(run)) throw Error('Use a versioned compatibility evidence directory.');
async function stored(page: Page, name: string) {
  return page.evaluate(async name => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const r = indexedDB.open('noffice-workspace',1);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);
    });
    try { return await new Promise<OfficeFile>((resolve,reject)=>{
      const r=db.transaction('files').objectStore('files').getAll();
      r.onsuccess=()=>resolve(r.result.find((f: OfficeFile)=>f.name===name));r.onerror=()=>reject(r.error);
    }); } finally { db.close(); }
  },name);
}
async function removeMetadata(page: Page, name: string, brokenOriginal?: string) {
  return page.evaluate(async ({name,brokenOriginal})=>{
    const db=await new Promise<IDBDatabase>((resolve,reject)=>{const r=indexedDB.open('noffice-workspace',1);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
    try { return await new Promise<number>((resolve,reject)=>{
      const tx=db.transaction('files','readwrite'),store=tx.objectStore('files'),request=store.getAll();let revision=0;
      request.onsuccess=()=>{
        const file=request.result.find((f: OfficeFile)=>f.name===name) as OfficeFile;
        if(file.content.kind!=='word')throw Error();
        delete file.content.docxStructure!.compatibility;
        if(brokenOriginal)file.original!.data=Uint8Array.from(atob(brokenOriginal),c=>c.charCodeAt(0)).buffer;
        revision=file.revision;store.put(file);
      };
      tx.oncomplete=()=>resolve(revision);tx.onabort=()=>reject(tx.error);tx.onerror=()=>reject(tx.error);
    }); }finally{db.close();}
  },{name,brokenOriginal});
}
const reopen = async (page: Page,name:string) => {
  await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();
  await page.getByRole('button',{name,exact:true}).click();
};
async function header(page: Page) {
  await page.getByRole('button',{name:'Insert',exact:true}).click();
  await page.getByRole('button',{name:'Headers and footers',exact:true}).click();
  await page.getByRole('button',{name:/^Section 1 \u2014 Default header/}).click();
  return page.getByRole('textbox',{name:'Header or footer text',exact:true});
}
async function download(page:Page,path:string,label:string|RegExp,recovery=false){
  if(!recovery)await page.getByRole('button',{name:'Export',exact:true}).click();
  const waiting=page.waitForEvent('download');await page.getByRole('button',{name:label,exact:typeof label==='string'}).click();
  await(await waiting).saveAs(path);return hash(await fs.readFile(path));
}
for(const sample of ['mode-14','mode-15','renamed-mode15','orphan-settings','settings-missing','mode-16'])
test(`Word compatibility ${sample}: real edits, history, migration, reload and retained exports`,async({page})=>{
  test.setTimeout(120000);
  const source=await fs.readFile(`tests/fixtures/word-compatibility-${sample}.docx`), reference=native.rows.find(r=>r.name===sample)!;
  expect(hash(source)).toBe(reference.inputHash);
  const name='Compatibility '+sample,root=`.local/word-compatibility/${run}/${sample}`;await fs.mkdir(root,{recursive:true});
  const errors:string[]=[];page.on('pageerror',e=>errors.push(e.message));await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles({name:name+'.docx',buffer:source,mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const body=page.getByRole('textbox',{name:'Document text',exact:true});await expect(body).toBeVisible();
  const outputs:Record<string,string>={};
  outputs['original.docx']=await download(page,root+'/original.docx','DOCX file Editable in Microsoft Word');expect(outputs['original.docx']).toBe(hash(source));
  await page.getByRole('button',{name:'Home',exact:true}).click();await body.focus();await page.keyboard.press('Control+Home');await page.keyboard.press('End');await page.keyboard.type(' edited');
  await expect(body.locator('p').first()).toHaveText('Body1 edited');
  await page.getByRole('button',{name:'Undo',exact:true}).click();await expect(body.locator('p').first()).toHaveText('Body1');
  await page.getByRole('button',{name:'Redo',exact:true}).click();await expect(body.locator('p').first()).toHaveText('Body1 edited');
  const editor=await header(page);await editor.focus();await page.keyboard.press('Control+Home');await page.keyboard.press('Control+ArrowDown');await page.keyboard.press('Control+ArrowDown');await page.keyboard.press('Home');await page.keyboard.type('Saved ');
  await expect(editor.locator('p').nth(2)).toContainText('Saved Alpha');
  const dialog=page.getByRole('dialog');await dialog.getByRole('button',{name:'Undo',exact:true}).click();await expect(editor.locator('p').nth(2)).not.toContainText('Saved Alpha');
  await dialog.getByRole('button',{name:'Redo',exact:true}).click();await expect(editor.locator('p').nth(2)).toContainText('Saved Alpha');
  await dialog.getByRole('button',{name:'Apply header or footer',exact:true}).click();
  let even:boolean|undefined;
  if(sample==='renamed-mode15'){
    await page.getByRole('button',{name:'Headers and footers',exact:true}).click();await page.getByRole('button',{name:'Page options',exact:true}).click();
    const control=page.getByRole('checkbox',{name:'Different odd and even pages',exact:true});even=!(await control.isChecked());await control.setChecked(even);
    await page.getByRole('button',{name:'Apply page options',exact:true}).click();
    await page.getByRole('button',{name:'Home',exact:true}).click();await page.getByRole('button',{name:'Undo',exact:true}).click();await page.getByRole('button',{name:'Redo',exact:true}).click();
  }
  await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  outputs['edited.docx']=await download(page,root+'/edited.docx','DOCX file Editable in Microsoft Word');
  const before=await stored(page,name);expect(before.content.kind).toBe('word');if(before.content.kind!=='word')throw Error();
  const compatibility=before.content.docxStructure!.compatibility;expect(compatibility?.mode).toBe(sample==='mode-16'?null:['orphan-settings','settings-missing'].includes(sample)?12:reference.mode);
  const revision=await removeMetadata(page,name);await reopen(page,name);await expect(body.locator('p').first()).toHaveText('Body1 edited');
  await expect((await header(page)).locator('p').nth(2)).toContainText('Saved Alpha');await page.getByRole('button',{name:'Cancel',exact:true}).click();
  outputs['reloaded.noffice']=await download(page,root+'/reloaded.noffice',/^Noffice backup/);
  const backup=JSON.parse(await fs.readFile(root+'/reloaded.noffice','utf8'));expect(backup.content.docxStructure.compatibility).toEqual(compatibility);
  expect(backup.original.base64).toBe(source.toString('base64'));expect((await stored(page,name)).revision).toBe(revision);
  outputs['reloaded.docx']=await download(page,root+'/reloaded.docx','DOCX file Editable in Microsoft Word');
  const a=await JSZip.loadAsync(await fs.readFile(root+'/edited.docx')),b=await JSZip.loadAsync(await fs.readFile(root+'/reloaded.docx'));
  for(const entry of Object.values(a.files))if(!entry.dir)expect(await b.file(entry.name)!.async('uint8array'),entry.name).toEqual(await entry.async('uint8array'));
  if(sample==='renamed-mode15'){expect(a.file('word/settings.xml')).toBeNull();expect(a.file('word/metadata/settings-custom.xml')).not.toBeNull();}
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles(root+'/edited.docx');await expect(body.locator('p').first()).toHaveText('Body1 edited');
  await expect((await header(page)).locator('p').nth(2)).toContainText('Saved Alpha');await page.getByRole('button',{name:'Cancel',exact:true}).click();expect(errors).toEqual([]);
  if(even!==undefined){
    await page.getByRole('button',{name:'Headers and footers',exact:true}).click();await page.getByRole('button',{name:'Page options',exact:true}).click();
    await expect(page.getByRole('checkbox',{name:'Different odd and even pages',exact:true})).toBeChecked({checked:even});
    await page.getByRole('button',{name:'Cancel',exact:true}).click();
  }
  await fs.writeFile(root+'/browser-report.json',JSON.stringify({sample,sourceHash:hash(source),outputs,compatibility,revision,actions:{bodySuffix:' edited',headerPrefix:'Saved ',even},errors,buildHash:await wordStoryBuildHash(),testHash:hash(await fs.readFile('tests/word-compatibility.spec.ts'))},null,2));
});

test('Word compatibility broken legacy settings preserve saved work and original recovery',async({page})=>{
  const source=await fs.readFile('tests/fixtures/word-compatibility-renamed-mode15.docx'),zip=await JSZip.loadAsync(source);zip.remove('word/metadata/settings-custom.xml');
  const damaged=await zip.generateAsync({type:'nodebuffer'}),name='Compatibility damaged',root=`.local/word-compatibility/${run}/damaged`;await fs.mkdir(root,{recursive:true});
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles({name:name+'.docx',buffer:source,mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const body=page.getByRole('textbox',{name:'Document text',exact:true});await expect(body).toBeVisible();await body.focus();await page.keyboard.press('Control+Home');await page.keyboard.type('Saved ');
  await expect(body.locator('p').first()).toHaveText('Saved Body1');await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  await removeMetadata(page,name,damaged.toString('base64'));const before=await stored(page,name);await reopen(page,name);
  await expect(page.getByRole('region',{name:'Document recovery'}).getByRole('alert')).toContainText('settings are missing');await expect(body).toHaveCount(0);
  const backupHash=await download(page,root+'/recovered.noffice','Export Noffice backup',true),originalHash=await download(page,root+'/original.docx','Download original DOCX',true);
  const backup=JSON.parse(await fs.readFile(root+'/recovered.noffice','utf8'));expect(backup.content).toEqual(before.content);expect(backup.original.base64).toBe(damaged.toString('base64'));expect(originalHash).toBe(hash(damaged));expect((await stored(page,name)).revision).toBe(before.revision);
  await fs.writeFile(root+'/browser-report.json',JSON.stringify({rejected:true,backupHash,originalHash,revision:before.revision,buildHash:await wordStoryBuildHash(),testHash:hash(await fs.readFile('tests/word-compatibility.spec.ts'))},null,2));
});
