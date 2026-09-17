import {test,expect} from '@playwright/test';
import fs from 'node:fs/promises';
import {createHash} from 'node:crypto';
import reference from './fixtures/native-word-script-fonts.json' with {type:'json'};
import {wordStoryBuildHash} from './word-story-artifacts';

for(const row of reference.rows)test(`Word scripts ${row.name}: commands, history, files and recovery`,async({page})=>{
  test.setTimeout(60000);
  const run=process.env.NOFFICE_SCRIPT_FONT_RUN||'browser-v1';if(!/^browser-v\d+$/.test(run))throw Error('Invalid evidence folder');
  const root=`.local/word-script-fonts/${run}/${row.name}`;await fs.mkdir(root,{recursive:true});
  const hash=(bytes:Buffer)=>createHash('sha256').update(bytes).digest('hex'),source=await fs.readFile('tests/fixtures/'+row.sourceFile);
  expect(hash(source)).toBe(row.stages[0].docxHash);const body=row.kind==='Body',kind=row.kind==='Headers'?'header':'footer';
  const other=row.script==='Superscript'?'Subscript':'Superscript',opposite=reference.rows.find(r=>r.kind===row.kind&&r.size===row.size&&r.script===other)!;
  const name='Script fonts '+row.name,errors:string[]=[];page.on('pageerror',error=>errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({name:name+'.docx',buffer:source,mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const open=async()=>{
    if(body)return;await page.getByRole('button',{name:'Insert',exact:true}).click();await page.getByRole('button',{name:'Headers and footers',exact:true}).click();
    await page.getByRole('button',{name:new RegExp(`^Section 1 — Default ${kind}`)}).click();
  };
  const editor=page.getByRole('textbox',{name:body?'Document text':'Header or footer text',exact:true}),controls=body?page:page.getByRole('dialog').last();
  const snapshot=()=>editor.evaluate(element=>{
    type Node={type:string;text?:string;attrs:Record<string,unknown>;content?:Node[];marks?:{type:string;attrs:Record<string,unknown>}[]};
    const p=(element as HTMLElement&{editor:{getJSON():Node}}).editor.getJSON().content![0];
    const read=(text:string,marks:Node['marks']=[])=>{
      const attrs=marks?.find(m=>m.type==='textStyle')?.attrs;
      return {text,family:String(attrs?.fontFamily||p.attrs.paragraphFontFamily).replace(/^["']|["']$/g,''),size:parseFloat(String(attrs?.fontSize||p.attrs.paragraphFontSize)),
        superscript:marks?.some(m=>m.type==='superscript')?-1:0,subscript:marks?.some(m=>m.type==='subscript')?-1:0};
    };
    return [...p.content!.flatMap(n=>[...(n.type==='wordTab'?'\t':n.text||'')].map(c=>read(c,n.marks))),read('\r')];
  });
  const selection=()=>editor.evaluate(element=>{
    const e=(element as HTMLElement&{editor:{state:{selection:{from:number;to:number}}}}).editor;
    return {from:e.state.selection.from,to:e.state.selection.to};
  });
  const select=async()=>{
    await editor.focus();await page.keyboard.press('Control+Home');await expect.poll(selection).toEqual({from:1,to:1});
    await page.keyboard.press('ArrowRight');await expect.poll(selection).toEqual({from:2,to:2});
    await page.keyboard.press('ArrowRight');await expect.poll(selection).toEqual({from:3,to:3});
    await page.keyboard.press('Shift+ArrowRight');await expect.poll(selection).toEqual({from:3,to:4});
  };
  const click=async(label:string)=>controls.getByRole('button',{name:label,exact:true}).click();
  await open();await expect.poll(snapshot).toEqual(row.stages[0].characters);await select();
  await expect(controls.getByRole('button',{name:row.script,exact:true})).toBeVisible();await click(row.script);
  await expect.poll(snapshot).toEqual(row.stages[1].characters);
  await expect.poll(() => editor.evaluate(element => {
    const p = element.querySelector('p')!, walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) if (walker.currentNode.textContent === 'B')
      return parseFloat(getComputedStyle(walker.currentNode.parentElement!).fontSize) * .75;
    return null;
  })).toBeCloseTo(row.geometry.b.size, 3);
  await expect(controls.getByRole('spinbutton',{name:'Font size',exact:true})).toHaveValue(String(row.size));
  await click('Undo');expect(await snapshot()).toEqual(row.stages[0].characters);await click('Redo');expect(await snapshot()).toEqual(row.stages[1].characters);
  await click(other);expect(await snapshot()).toEqual(opposite.stages[1].characters);await click('Undo');expect(await snapshot()).toEqual(row.stages[1].characters);
  await click(row.script);expect(await snapshot()).toEqual(row.stages[0].characters);await click('Undo');expect(await snapshot()).toEqual(row.stages[1].characters);
  await editor.focus();await page.keyboard.insertText('N');await expect(editor.locator('p').first()).toHaveText('A\tN');
  await page.keyboard.press('Control+z');expect(await snapshot()).toEqual(row.stages[1].characters);
  await page.keyboard.press('Control+y');await expect(editor.locator('p').first()).toHaveText('A\tN');await page.keyboard.press('Control+z');
  expect(await snapshot()).toEqual(row.stages[1].characters);
  if(!body){
    await page.getByRole('button',{name:'Cancel',exact:true}).click();await open();expect(await snapshot()).toEqual(row.stages[0].characters);
    await select();await click(row.script);await page.getByRole('button',{name:'Apply header or footer',exact:true}).click();
    await page.getByRole('button',{name:'Home',exact:true}).click();await page.getByRole('button',{name:'Undo',exact:true}).click();
    await open();expect(await snapshot()).toEqual(row.stages[0].characters);await page.getByRole('button',{name:'Cancel',exact:true}).click();
    await page.getByRole('button',{name:'Home',exact:true}).click();await page.getByRole('button',{name:'Redo',exact:true}).click();
  }
  const outputs:Record<string,string>={};
  const download=async(file:string,label:string)=>{
    await page.getByRole('button',{name:'Export',exact:true}).click();const pending=page.waitForEvent('download');
    await page.getByRole('button',{name:label,exact:true}).click();await(await pending).saveAs(root+'/'+file);outputs[file]=hash(await fs.readFile(root+'/'+file));
  };
  await download('edited.docx','DOCX file Editable in Microsoft Word');await download('edited.pdf','PDF file');
  await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();await page.getByRole('button',{name,exact:true}).click();
  await download('reloaded.pdf','PDF file');await open();expect(await snapshot()).toEqual(row.stages[1].characters);
  await select();const size=controls.getByRole('spinbutton',{name:'Font size',exact:true});await size.fill('0');await size.press('Tab');
  expect(await snapshot()).toEqual(row.stages[1].characters);await expect(size).toHaveValue(String(row.size));
  await select();await size.fill('12');await size.press('Enter');expect((await snapshot())[2].size).toBe(12);
  if(!body)await page.getByRole('button',{name:'Apply header or footer',exact:true}).click();
  const refused:string[]=[];const observe=(download:{suggestedFilename:()=>string})=>refused.push(download.suggestedFilename());page.on('download',observe);
  await page.getByRole('button',{name:'Export',exact:true}).click();await page.getByRole('button',{name:'PDF file',exact:true}).click();
  await expect(page.locator('.error-banner')).toContainText('Export failed:');expect(refused).toEqual([]);page.off('download',observe);
  await download('unsupported.noffice','Noffice backup Full editable model and retained original · .noffice');
  const backup=JSON.parse((await fs.readFile(root+'/unsupported.noffice')).toString());expect(Buffer.from(backup.original.base64,'base64')).toEqual(source);
  await page.getByRole('button',{name:'Home',exact:true}).click();await page.getByRole('button',{name:'Undo',exact:true}).click();
  await download('recovered.pdf','PDF file');
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles(root+'/edited.docx');await open();expect(await snapshot()).toEqual(row.stages[1].characters);
  expect(errors).toEqual([]);await fs.writeFile(root+'/browser-report.json',JSON.stringify({outputs,errors,sourceHash:hash(source),buildHash:await wordStoryBuildHash(),
    testHash:hash(await fs.readFile('tests/word-script-fonts.spec.ts')),referenceHash:hash(await fs.readFile('tests/fixtures/native-word-script-fonts.json'))},null,2));
});
