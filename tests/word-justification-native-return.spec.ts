import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import arial from './fixtures/native-word-justification-browser.json' with { type:'json' };
import calibri from './fixtures/native-word-justification-calibri-browser.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
const profile = process.env.NOFFICE_JUSTIFICATION_PROFILE || 'arial-ten';
if (!['arial-ten', 'calibri-eleven'].includes(profile)) throw Error('Invalid justification profile');
const reference = profile === 'calibri-eleven' ? calibri : arial;
const hash=(bytes:Buffer)=>createHash('sha256').update(bytes).digest('hex');
const sourceRun=process.env.NOFFICE_JUSTIFICATION_SOURCE||'browser-v7';
const run=process.env.NOFFICE_JUSTIFICATION_RETURN||'returns-browser-v1';
if(!/^browser-v\d+$/.test(sourceRun)||!/^returns-browser-v\d+$/.test(run))throw Error('Versioned justification evidence required.');

for(const row of reference.rows) test(`Word justification native return ${row.name}`,async({page})=>{
  const sourceRoot=`.local/word-story-paragraph-layout/justification-${profile}/${sourceRun}`;
  const native=JSON.parse((await fs.readFile(sourceRoot+'/native-report.json','utf8')).replace(/^\uFEFF/,''));
  const state=native.rows.find((r:{name:string})=>r.name===row.name).states.find((s:{stage:string})=>s.stage==='edited-returned');
  const sourcePath=sourceRoot+'/'+row.name+'/native-v1/edited-returned.docx',source=await fs.readFile(sourcePath);
  expect(hash(source)).toBe(state.inputHash);
  const root=`.local/word-story-paragraph-layout/justification-${profile}/${run}/${row.name}`;
  await fs.mkdir(root,{recursive:true});
  const name='Native justification '+row.name,original='New '+row.text+' Native',final='More '+original;
  const errors:string[]=[];page.on('pageerror',error=>errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({name:name+'.docx',buffer:source,mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const target=row.kind==='body'?page.getByRole('textbox',{name:'Document text',exact:true}).locator('p').first()
    :page.locator('.word-story-measure-host p').filter({hasText:row.text}).first();
  await expect(page.locator('.section-page')).toHaveCount(state.state.pages);await expect(target).toHaveText(original);
  const outputs:Record<string,string>={};
  const download=async(file:string,label:string|RegExp)=>{
    await page.getByRole('button',{name:'Export',exact:true}).click();const waiting=page.waitForEvent('download');
    await page.getByRole('button',{name:label,exact:typeof label==='string'}).click();
    await(await waiting).saveAs(root+'/'+file);outputs[file]=hash(await fs.readFile(root+'/'+file));
  };
  await download('original.docx','DOCX file Editable in Microsoft Word');expect(outputs['original.docx']).toBe(hash(source));
  let editor=page.getByRole('textbox',{name:'Document text',exact:true});
  if(row.kind!=='body'){
    await page.getByRole('button',{name:'Insert',exact:true}).click();await page.getByRole('button',{name:'Headers and footers',exact:true}).click();
    await page.getByRole('button',{name:new RegExp('^Section 1 \\u2014 Default '+row.kind)}).click();
    editor=page.getByRole('textbox',{name:'Header or footer text',exact:true});
  }
  const paragraph=editor.locator('p').filter({hasText:row.text}).first();
  await editor.focus();await page.keyboard.press('Control+Home');
  if(row.kind!=='body'){await page.keyboard.press('Control+ArrowDown');await page.keyboard.press('Control+ArrowDown');await page.keyboard.press('Home');}
  await page.keyboard.type('More ');await expect(paragraph).toHaveText(final);
  const toolbar=row.kind==='body'?page:page.getByRole('dialog');
  await toolbar.getByRole('button',{name:'Undo',exact:true}).click();await expect(paragraph).toHaveText(original);
  await toolbar.getByRole('button',{name:'Redo',exact:true}).click();await expect(paragraph).toHaveText(final);
  if(row.kind!=='body')await page.getByRole('button',{name:'Apply header or footer',exact:true}).click();
  await expect(target).toHaveText(final);await expect(target).toHaveAttribute('data-word-justification','modern');
  await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();await page.getByRole('button',{name,exact:true}).click();
  await expect(target).toHaveText(final);
  await download('edited.docx','DOCX file Editable in Microsoft Word');await download('edited.pdf','PDF file');
  await download('reloaded.noffice',/^Noffice backup/);
  const backup=JSON.parse(await fs.readFile(root+'/reloaded.noffice','utf8'));
  expect(backup.original.base64).toBe(source.toString('base64'));
  expect(JSON.stringify(backup.content)).not.toMatch(/data-word-justification|word-spacing:|wordJustificationBreak/);
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles(root+'/edited.docx');
  await expect(target).toHaveText(final);await expect(target).toHaveAttribute('data-word-justification','modern');expect(errors).toEqual([]);
  await fs.writeFile(root+'/browser-report.json',JSON.stringify({name:row.name,kind:row.kind,sourcePath,sourceHash:hash(source),outputs,actions:{prefix:'More '},errors,
    nativeReceiptHash:hash(await fs.readFile(sourceRoot+'/native-report.json')),buildHash:await wordStoryBuildHash(),testHash:hash(await fs.readFile('tests/word-justification-native-return.spec.ts'))},null,2));
});
