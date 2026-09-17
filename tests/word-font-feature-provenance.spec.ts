import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import type { OfficeFile } from '../src/model';
import reference from './fixtures/native-word-font-feature-provenance.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const row of reference.rows) test(`Word font-feature provenance ${row.name}: real authoring and legacy recovery`, async ({ page }) => {
  const run = process.env.NOFFICE_FONT_PROVENANCE_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
  const root = `.local/word-font-feature-provenance/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const source = await fs.readFile('tests/fixtures/'+row.sourceFile);
  expect(hash(source)).toBe(row.stages[0].docxHash);
  const name = 'Font provenance '+row.name, kind = row.kind === 'Headers' ? 'header' : 'footer';
  const expected = row.stages[1].state.text.replace(/\r$/, '');
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({ name:name+'.docx', buffer:source,
    mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const menu = async () => {
    await page.getByRole('button',{name:'Insert',exact:true}).click();
    await page.getByRole('button',{name:'Headers and footers',exact:true}).click();
  };
  const openStory = async () => {
    await menu();await page.getByRole('button',{name:new RegExp(`^Section ${row.section} — Default ${kind}`)}).click();
  };
  if (row.mode === 'cloned') {
    await menu();await page.getByRole('button',{name:'Section links',exact:true}).click();
    await page.getByLabel('Header or footer',{exact:true}).selectOption(kind);
    await page.getByLabel('Header/footer page type',{exact:true}).selectOption('default');
    await page.getByLabel('Link section',{exact:true}).selectOption({label:'Section 2'});
    await page.getByRole('checkbox',{name:'Link to previous',exact:true}).uncheck();
    await page.getByRole('button',{name:'Apply header/footer link',exact:true}).click();
  }
  await openStory();
  const editor=page.getByRole('textbox',{name:'Header or footer text',exact:true});
  const initial=await editor.textContent();await editor.focus();await page.keyboard.press('Control+End');
  await page.keyboard.insertText(row.mode==='cloned'?' New':'Native');
  await expect(editor).toHaveText(expected);
  await page.keyboard.press('Control+z');await expect(editor).toHaveText(initial!);
  await page.keyboard.press('Control+y');await expect(editor).toHaveText(expected);
  await page.getByRole('button',{name:'Apply header or footer',exact:true}).click();
  await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  const stored=(strip:boolean)=>page.evaluate(({name,strip,mode,kind})=>new Promise<OfficeFile>((resolve,reject)=>{
    const request=indexedDB.open('noffice-workspace');request.onerror=()=>reject(request.error);
    request.onsuccess=()=>{
      const db=request.result,tx=db.transaction('files',strip?'readwrite':'readonly'),store=tx.objectStore('files'),all=store.getAll();let result:OfficeFile;
      all.onsuccess=()=>{
        result=all.result.find((file:OfficeFile)=>file.name===name);
        if(!result||result.content.kind!=='word'||!result.content.stories){tx.abort();return;}
        const target=result.content.stories.parts.find(part=>part.kind===kind&&(mode==='cloned'?!!part.copiedFrom:part.created));
        if(!target){tx.abort();return;}
        if(strip){
          const clean=(html:string)=>{
            const doc=new DOMParser().parseFromString(html,'text/html');
            for(const element of doc.querySelectorAll<HTMLElement>('[data-word-font-features],[data-word-paragraph-font-features]')){
              element.removeAttribute('data-word-font-features');element.removeAttribute('data-word-paragraph-font-features');element.style.removeProperty('font-feature-settings');
            }
            return doc.body.innerHTML;
          };
          delete result.content.fontFeaturesVersion;result.content.html=clean(result.content.html);
          for(const part of result.content.stories.parts)part.html=clean(part.html);
          store.put(result);
        }
      };
      tx.oncomplete=()=>{db.close();resolve(result);};tx.onerror=tx.onabort=()=>{db.close();reject(tx.error||Error('Missing actual authored provenance'));};
    };
  }),{name,strip,mode:row.mode,kind});
  const authored=await stored(false);expect(authored.content.kind).toBe('word');
  const legacy=await stored(true);
  await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();
  await page.getByRole('button',{name,exact:true}).click();await openStory();await expect(editor).toHaveText(expected);
  const features=()=>editor.evaluate(element=>{
    const model=(element as HTMLElement&{editor:{getJSON():{content:{attrs:Record<string,unknown>;content?:{marks?:{type:string;attrs:Record<string,unknown>}[]}[]}[]}}}).editor.getJSON();
    return model.content.map(p=>({mark:p.attrs.paragraphFontFeatures,runs:[...new Set((p.content||[]).map(n=>n.marks?.find(m=>m.type==='textStyle')?.attrs.wordFontFeatures))]}));
  });
  expect(await features()).toEqual([{mark:1,runs:[1]}]);
  expect(await stored(false)).toEqual(legacy);
  await editor.focus();await page.keyboard.press('Control+End');await page.keyboard.insertText('Z');
  await expect(editor).toHaveText(expected+'Z');await page.keyboard.press('Control+z');await expect(editor).toHaveText(expected);
  await page.keyboard.press('Control+y');await expect(editor).toHaveText(expected+'Z');await page.keyboard.press('Control+z');
  await page.getByRole('button',{name:'Cancel',exact:true}).click();expect(await stored(false)).toEqual(legacy);
  const outputs:Record<string,string>={};
  const download=async(label:string,file:string)=>{
    await page.getByRole('button',{name:'Export',exact:true}).click();const pending=page.waitForEvent('download');
    await page.getByRole('button',{name:label,exact:true}).click();await(await pending).saveAs(root+'/'+file);
    const bytes=await fs.readFile(root+'/'+file);outputs[file]=hash(bytes);return bytes;
  };
  const backup=JSON.parse((await download('Noffice backup Full editable model and retained original · .noffice','recovered.noffice')).toString());
  expect(Buffer.from(backup.original.base64,'base64')).toEqual(source);expect(backup.content.fontFeaturesVersion).toBe(1);
  await download('DOCX file Editable in Microsoft Word','recovered.docx');await download('PDF file','recovered.pdf');
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles(root+'/recovered.docx');
  await openStory();await expect(editor).toHaveText(expected);expect(await features()).toEqual([{mark:1,runs:[1]}]);
  await page.getByRole('button',{name:'Cancel',exact:true}).click();await download('PDF file','reimported.pdf');
  expect(errors).toEqual([]);
  await fs.writeFile(root+'/browser-report.json',JSON.stringify({outputs,errors,sourceHash:hash(source),buildHash:await wordStoryBuildHash(),
    testHash:hash(await fs.readFile('tests/word-font-feature-provenance.spec.ts')),referenceHash:hash(await fs.readFile('tests/fixtures/native-word-font-feature-provenance.json'))},null,2));
});
