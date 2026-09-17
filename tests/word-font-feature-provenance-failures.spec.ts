import {test,expect} from '@playwright/test';
import fs from 'node:fs/promises';
import {createHash} from 'node:crypto';
import type {OfficeFile} from '../src/model';
import reference from './fixtures/native-word-font-feature-mixed-provenance.json' with {type:'json'};
import {wordStoryBuildHash} from './word-story-artifacts';

for(const row of reference.rows)test(`Word copied-story mixed font features ${row.kind}: ambiguous edits retain recovery`,async({page})=>{
  const run=process.env.NOFFICE_FONT_PROVENANCE_FAILURE_RUN||'browser-v1';if(!/^browser-v\d+$/.test(run))throw Error('Invalid evidence folder');
  const root=`.local/word-font-feature-mixed-provenance/${run}/${row.name}`;await fs.mkdir(root,{recursive:true});
  const hash=(bytes:Buffer)=>createHash('sha256').update(bytes).digest('hex'),source=await fs.readFile('tests/fixtures/'+row.sourceFile);
  expect(hash(source)).toBe(row.stages[0].docxHash);const kind=row.kind==='Headers'?'header':'footer',name='Ambiguous copied '+kind;
  const errors:string[]=[];page.on('pageerror',error=>errors.push(error.message));await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({name:name+'.docx',buffer:source,mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const menu=async()=>{await page.getByRole('button',{name:'Insert',exact:true}).click();await page.getByRole('button',{name:'Headers and footers',exact:true}).click();};
  await menu();await page.getByRole('button',{name:'Section links',exact:true}).click();
  await page.getByLabel('Header or footer',{exact:true}).selectOption(kind);await page.getByLabel('Header/footer page type',{exact:true}).selectOption('default');
  await page.getByLabel('Link section',{exact:true}).selectOption({label:'Section 2'});await page.getByRole('checkbox',{name:'Link to previous',exact:true}).uncheck();
  await page.getByRole('button',{name:'Apply header/footer link',exact:true}).click();
  await menu();await page.getByRole('button',{name:new RegExp(`^Section 2 — Default ${kind}`)}).click();
  const editor=page.getByRole('textbox',{name:'Header or footer text',exact:true});
  const flags=await editor.evaluate(element=>{
    const model=(element as HTMLElement&{editor:{getJSON():{content:{content:{marks:{type:string;attrs:Record<string,unknown>}[]}[]}[]}}}).editor.getJSON();
    return [...new Set(model.content[0].content.map(n=>n.marks.find(m=>m.type==='textStyle')!.attrs.wordFontFeatures))].sort();
  });expect(flags).toEqual([0,1]);
  await editor.focus();await page.keyboard.press('Control+End');await page.keyboard.insertText(' New');
  const expected=row.stages[1].state.text.replace(/\r$/,'');await expect(editor).toHaveText(expected);
  await page.keyboard.press('Control+z');await expect(editor).toHaveText(row.stages[0].state.text.replace(/\r$/,''));
  await page.keyboard.press('Control+y');await expect(editor).toHaveText(expected);await page.getByRole('button',{name:'Apply header or footer',exact:true}).click();
  await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  const stored=(strip:boolean)=>page.evaluate(({name,kind,strip})=>new Promise<OfficeFile>((resolve,reject)=>{
    const request=indexedDB.open('noffice-workspace');request.onerror=()=>reject(request.error);
    request.onsuccess=()=>{
      const db=request.result,tx=db.transaction('files',strip?'readwrite':'readonly'),store=tx.objectStore('files'),all=store.getAll();let file:OfficeFile;
      all.onsuccess=()=>{
        file=all.result.find((f:OfficeFile)=>f.name===name);
        if(!file||file.content.kind!=='word'||!file.content.stories?.parts.some(p=>p.kind===kind&&p.copiedFrom)){tx.abort();return;}
        if(strip){
          const clean=(html:string)=>{
            const dom=new DOMParser().parseFromString(html,'text/html');
            for(const el of dom.querySelectorAll<HTMLElement>('[data-word-font-features],[data-word-paragraph-font-features]')){
              el.removeAttribute('data-word-font-features');el.removeAttribute('data-word-paragraph-font-features');el.style.removeProperty('font-feature-settings');
            }return dom.body.innerHTML;
          };
          delete file.content.fontFeaturesVersion;file.content.html=clean(file.content.html);
          for(const part of file.content.stories!.parts)part.html=clean(part.html);store.put(file);
        }
      };
      tx.oncomplete=()=>{db.close();resolve(file);};tx.onerror=tx.onabort=()=>{db.close();reject(tx.error||Error('Missing copied provenance'));};
    };
  }),{name,kind,strip});
  const legacy=await stored(true);await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();await page.getByRole('button',{name,exact:true}).click();
  await expect(page.getByRole('region',{name:'Document recovery'}).getByRole('alert')).toHaveText(/edited text with missing mixed font-feature metadata/);
  await expect(page.getByRole('textbox',{name:'Document text',exact:true})).toHaveCount(0);expect(await stored(false)).toEqual(legacy);
  const downloads:string[]=[];page.on('download',download=>downloads.push(download.suggestedFilename()));
  for(const label of ['DOCX file Editable in Microsoft Word','PDF file']){
    await page.getByRole('button',{name:'Export',exact:true}).click();await page.getByRole('button',{name:label,exact:true}).click();
    await expect(page.locator('.error-banner')).toContainText(/Export failed:.*edited text with missing mixed font-feature metadata/);
    expect(downloads).toEqual([]);expect(await stored(false)).toEqual(legacy);
    await page.getByRole('button',{name:'Dismiss error',exact:true}).click();
  }
  await page.evaluate(()=>{window.print=()=>{throw Error('Blocked recovery must not open print');};});
  await page.getByRole('button',{name:'Export',exact:true}).click();await page.getByRole('button',{name:'Print / Save as PDF',exact:true}).click();
  await expect(page.locator('.error-banner')).toContainText(/Print unavailable:.*edited text with missing mixed font-feature metadata/);
  expect(downloads).toEqual([]);expect(await stored(false)).toEqual(legacy);await page.getByRole('button',{name:'Dismiss error',exact:true}).click();
  const originalPending=page.waitForEvent('download');await page.getByRole('button',{name:'Download original DOCX',exact:true}).click();
  await(await originalPending).saveAs(root+'/original.docx');expect(await fs.readFile(root+'/original.docx')).toEqual(source);
  await page.getByRole('button',{name:'Export',exact:true}).click();const pending=page.waitForEvent('download');
  await page.getByRole('button',{name:'Noffice backup Full editable model and retained original · .noffice',exact:true}).click();await(await pending).saveAs(root+'/recovery.noffice');
  const backupBytes=await fs.readFile(root+'/recovery.noffice'),backup=JSON.parse(backupBytes.toString());
  expect(backup.content).toEqual(legacy.content);expect(Buffer.from(backup.original.base64,'base64')).toEqual(source);expect(await stored(false)).toEqual(legacy);
  await page.locator('input[type=file][multiple]').setInputFiles(root+'/original.docx');
  await expect(page.getByRole('textbox',{name:'Document text',exact:true})).toBeVisible();
  await expect(page.getByRole('region',{name:'Document recovery'})).toHaveCount(0);
  await page.getByRole('button',{name:'Export',exact:true}).click();const restoredPending=page.waitForEvent('download');
  await page.getByRole('button',{name:'DOCX file Editable in Microsoft Word',exact:true}).click();await(await restoredPending).saveAs(root+'/reopened-original.docx');
  expect(await fs.readFile(root+'/reopened-original.docx')).toEqual(source);expect(await stored(false)).toEqual(legacy);
  expect(errors).toEqual([]);await fs.writeFile(root+'/browser-report.json',JSON.stringify({sourceHash:hash(source),backupHash:hash(backupBytes),originalHash:hash(await fs.readFile(root+'/original.docx')),errors,unchangedStorage:true,
    buildHash:await wordStoryBuildHash(),testHash:hash(await fs.readFile('tests/word-font-feature-provenance-failures.spec.ts')),
    referenceHash:hash(await fs.readFile('tests/fixtures/native-word-font-feature-mixed-provenance.json'))},null,2));
});
