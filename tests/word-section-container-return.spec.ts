import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/word-section-containers/reference.json' with { type: 'json' };
import freshReference from './fixtures/word-section-containers/fresh-reference.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

const nativeRun=process.env.NOFFICE_SECTION_CONTAINER_NATIVE;
const hash=(data:Buffer)=>createHash('sha256').update(data).digest('hex');
for(const sample of [...reference.rows.filter(r=>!r.name.startsWith('table')), ...(nativeRun?.startsWith('fresh-')?freshReference.rows:[])])
  test(`Native list return ${sample.name}: edit, history, reload, preserve original`,async({page})=>{
    test.skip(!nativeRun,'Requires the independently saved native Word return files.');
    if(!/^(fresh-)?exports-native-v\d+$/.test(nativeRun!))throw Error('Invalid native return run');
    const run=process.env.NOFFICE_SECTION_CONTAINER_RETURN || 'returns-v1';
    if(!/^returns-v\d+$/.test(run))throw Error('Invalid browser return run');
    const sourcePath=`.local/word-section-containers/${nativeRun}/${sample.name}/edited-return.docx`;
    const root=`.local/word-section-containers/${nativeRun!.startsWith('fresh-')?'fresh-':''}${run}/${sample.name}`;await fs.mkdir(root,{recursive:true});
    const source=await fs.readFile(sourcePath);
    await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles(sourcePath);
    const editor=page.getByRole('textbox',{name:'Document text',exact:true});await expect(editor).toBeVisible();
    const model=()=>editor.evaluate(el=>(el as HTMLElement & {editor:import('@tiptap/core').Editor}).editor.getJSON());
    const before=await model();
    await editor.evaluate(el=>{
      const e=(el as HTMLElement & {editor:import('@tiptap/core').Editor}).editor;
      let position:number|undefined;
      e.state.doc.descendants((n,p)=>{if(n.isTextblock){if(n.textContent.includes('X'))position=p+2+n.textContent.indexOf('X');return false;}});
      if(position===undefined)throw Error('Native edit missing');e.commands.setTextSelection(position);e.view.focus();
    });
    await page.keyboard.insertText('Y');const edited=await model();expect(edited).not.toEqual(before);
    await page.keyboard.press('Control+z');await expect.poll(model).toEqual(before);
    await page.keyboard.press('Control+y');await expect.poll(model).toEqual(edited);
    await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
    const name=await page.getByRole('textbox',{name:'File name',exact:true}).inputValue();
    await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();
    await page.getByRole('button',{name,exact:true}).click();await expect(editor).toBeVisible();
    expect(await model()).toEqual(edited);
    const stored=await page.evaluate(async()=>{
      const db=await new Promise<IDBDatabase>((resolve,reject)=>{const r=indexedDB.open('noffice-workspace',1);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
      try{const files=await new Promise<{original?:{data:ArrayBuffer}}[]>((resolve,reject)=>{const r=db.transaction('files').objectStore('files').getAll();r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
        return await Promise.all(files.map(async file=>file.original ? [...new Uint8Array(await crypto.subtle.digest('SHA-256',file.original.data))].map(b=>b.toString(16).padStart(2,'0')).join('') : null));}
      finally{db.close();}
    });
    // Export the edited file, then reopen the unchanged native source and prove
    // that no browser editing has mutated the retained original download.
    const capture=async(stage:string)=>{await page.getByRole('button',{name:'Export',exact:true}).click();const pending=page.waitForEvent('download');
      await page.getByRole('button',{name:'DOCX file Editable in Microsoft Word',exact:true}).click();const path=`${root}/${stage}.docx`;await(await pending).saveAs(path);return hash(await fs.readFile(path));};
    const editedHash=await capture('edited');
    await page.locator('input[type=file][multiple]').setInputFiles(sourcePath);await expect(editor).toBeVisible();
    const originalHash=await capture('original');expect(originalHash).toBe(hash(source));
    expect(stored.filter(Boolean)).toEqual([hash(source)]);
    await fs.writeFile(root+'/report.json',JSON.stringify({name:sample.name,sourceHash:hash(source),editedHash,originalHash,
      before,edited,buildHash:await wordStoryBuildHash(),testHash:hash(await fs.readFile('tests/word-section-container-return.spec.ts')),
      scope:'Native further edit reimport, real browser typing/history, edited reload/download, unchanged original download. Native comparison of the subsequent browser Y edit is separate.'},null,2));
  });
