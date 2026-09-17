import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';
import { wordStoryBuildHash } from './word-story-artifacts';

for (const mode of ['unsupported','missing','conflicting'])
test(`Fresh list ${mode}: failed export preserves editing, history, reload and backup`,async({page})=>{
  const run=process.env.NOFFICE_FRESH_LIST_RECOVERY || 'recovery-v1';
  if(!/^recovery-v\d+$/.test(run))throw Error('Invalid run');
  const root=`.local/word-section-containers/fresh-${run}/${mode}`;await fs.mkdir(root,{recursive:true});
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles('tests/fixtures/word-section-containers/numbered.docx');
  const editor=page.getByRole('textbox',{name:'Document text',exact:true});await expect(editor).toBeVisible();
  await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  const content=await page.evaluate(async()=>{
    const db=await new Promise<IDBDatabase>((resolve,reject)=>{const r=indexedDB.open('noffice-workspace',1);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
    try{return await new Promise<OfficeFile['content']>((resolve,reject)=>{const r=db.transaction('files').objectStore('files').getAll();r.onsuccess=()=>resolve(r.result.find((f:OfficeFile)=>f.name==='numbered').content);r.onerror=()=>reject(r.error);});}
    finally{db.close();}
  });
  if(content.kind!=='word')throw Error('Word required');
  if(mode==='missing')delete content.numbering;
  else if(mode==='unsupported')content.numbering!.paragraphs[0].numbering={status:'unsupported',reason:'Unqualified custom numbering'};
  else {const entry=content.numbering!.paragraphs[1].numbering;if(entry.status!=='resolved')throw Error('Resolved required');entry.definition.start=7;}
  const name=`Fresh list ${mode}`;
  await page.locator('input[type=file][multiple]').setInputFiles({name:name+'.noffice',mimeType:'application/json',
    buffer:Buffer.from(JSON.stringify({format:'noffice',version:1,name,content,warnings:[]}))});
  await expect(page.getByRole('textbox',{name:'File name',exact:true})).toHaveValue(name);
  const model=()=>editor.evaluate(el=>(el as HTMLElement & {editor:import('@tiptap/core').Editor}).editor.getJSON());
  const before=await model();
  await editor.evaluate(el=>{const e=(el as HTMLElement & {editor:import('@tiptap/core').Editor}).editor;e.commands.setTextSelection(2);e.view.focus();});
  await page.keyboard.insertText('X');const edited=await model();expect(edited).not.toEqual(before);
  await page.getByRole('button',{name:'Export',exact:true}).click();
  await page.getByRole('button',{name:'DOCX file Editable in Microsoft Word',exact:true}).click();
  await expect(page.getByRole('alert')).toContainText('retained DOCX numbering source');
  expect(await model()).toEqual(edited);
  await page.getByRole('button',{name:'Dismiss error',exact:true}).click();
  await editor.focus();await page.keyboard.press('Control+z');await expect.poll(model).toEqual(before);
  await page.keyboard.press('Control+y');await expect.poll(model).toEqual(edited);
  await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();
  await page.getByRole('button',{name,exact:true}).click();await expect(editor).toBeVisible();expect(await model()).toEqual(edited);
  await page.getByRole('button',{name:'Export',exact:true}).click();
  const pending=page.waitForEvent('download');await page.getByRole('button',{name:/Noffice backup/}).click();
  const path=`${root}/recovery.noffice`;await(await pending).saveAs(path);
  const saved=JSON.parse(await fs.readFile(path,'utf8'));
  expect(saved.content.numbering).toEqual(content.numbering);
  expect(saved.content.html).toBe(await editor.evaluate(el=>(el as HTMLElement & {editor:import('@tiptap/core').Editor}).editor.getHTML()));
  expect(await editor.textContent()).toContain('BXefore.');
  await fs.writeFile(root+'/report.json',JSON.stringify({mode,buildHash:await wordStoryBuildHash(),
    backupHash:createHash('sha256').update(await fs.readFile(path)).digest('hex'),
    testHash:createHash('sha256').update(await fs.readFile('tests/word-fresh-list-recovery.spec.ts')).digest('hex'),
    scope:'Source-derived fresh-package export rejects unsupported/missing/conflicting numbering. Actual typing/history/reload and backup preserve edits and metadata.'},null,2));
});
