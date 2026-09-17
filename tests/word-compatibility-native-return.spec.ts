import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
const hash=(bytes:Buffer)=>createHash('sha256').update(bytes).digest('hex');
const sourceRun=process.env.NOFFICE_COMPATIBILITY_SOURCE || 'browser-v1',run=process.env.NOFFICE_COMPATIBILITY_RETURN || 'returns-browser-v1';
if(!/^browser-v\d+$/.test(sourceRun)||!/^returns-browser-v\d+$/.test(run))throw Error('Versioned compatibility evidence required.');
for(const sample of ['mode-14','mode-15','renamed-mode15','orphan-settings','settings-missing','mode-16'])
test(`Word compatibility native ${sample} return retains edits, history, reload and original recovery`,async({page})=>{
  const sourceRoot=`.local/word-compatibility/${sourceRun}`,native=JSON.parse((await fs.readFile(sourceRoot+'/native-report.json','utf8')).replace(/^\uFEFF/,''));
  const row=native.rows.find((r:{name:string})=>r.name===sample),state=row.states.find((s:{stage:string})=>s.stage==='edited-returned');
  const sourcePath=sourceRoot+'/'+sample+'/native-v1/edited-returned.docx',source=await fs.readFile(sourcePath);expect(hash(source)).toBe(state.inputHash);
  const name='Native compatibility '+sample,root=`.local/word-compatibility/${run}/${sample}`;await fs.mkdir(root,{recursive:true});
  const errors:string[]=[];page.on('pageerror',e=>errors.push(e.message));await page.context().grantPermissions(['local-fonts']);
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles({name:name+'.docx',buffer:source,mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const body=page.getByRole('textbox',{name:'Document text',exact:true});await expect(body.locator('p').first()).toHaveText('Body1 edited Native');
  const outputs:Record<string,string>={};
  const download=async(file:string,label:string|RegExp)=>{
    await page.getByRole('button',{name:'Export',exact:true}).click();const waiting=page.waitForEvent('download');
    await page.getByRole('button',{name:label,exact:typeof label==='string'}).click();await(await waiting).saveAs(root+'/'+file);outputs[file]=hash(await fs.readFile(root+'/'+file));
  };
  await download('original.docx','DOCX file Editable in Microsoft Word');expect(outputs['original.docx']).toBe(hash(source));
  await page.getByRole('button',{name:'Home',exact:true}).click();await body.focus();await page.keyboard.press('Control+Home');await page.keyboard.press('End');await page.keyboard.type(' Browser');
  await expect(body.locator('p').first()).toHaveText('Body1 edited Native Browser');
  await page.getByRole('button',{name:'Undo',exact:true}).click();await expect(body.locator('p').first()).toHaveText('Body1 edited Native');
  await page.getByRole('button',{name:'Redo',exact:true}).click();await expect(body.locator('p').first()).toHaveText('Body1 edited Native Browser');
  await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();await page.getByRole('button',{name,exact:true}).click();
  await expect(body.locator('p').first()).toHaveText('Body1 edited Native Browser');
  await download('edited.docx','DOCX file Editable in Microsoft Word');await download('reloaded.noffice',/^Noffice backup/);
  const backup=JSON.parse(await fs.readFile(root+'/reloaded.noffice','utf8'));expect(backup.original.base64).toBe(source.toString('base64'));
  expect(backup.content.docxStructure.compatibility.mode).toBe(state.state.mode);expect(backup.content.docxStructure.compatibility.wordPerfectJustification).toBe(state.state.wpJustification);
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles(root+'/edited.docx');await expect(body.locator('p').first()).toHaveText('Body1 edited Native Browser');expect(errors).toEqual([]);
  await fs.writeFile(root+'/browser-report.json',JSON.stringify({sample,sourcePath,sourceHash:hash(source),outputs,bodySuffix:' Browser',errors,nativeReceiptHash:hash(await fs.readFile(sourceRoot+'/native-report.json')),buildHash:await wordStoryBuildHash(),testHash:hash(await fs.readFile('tests/word-compatibility-native-return.spec.ts'))},null,2));
});
