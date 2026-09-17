import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/word-section-containers/reference.json' with { type: 'json' };
import freshReference from './fixtures/word-section-containers/fresh-reference.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';
import { waitWordLayout, clickWordText } from './word-pagination-helpers';
import type { OfficeFile } from '../src/model';

const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
for (const sample of [...reference.rows.filter(r => !r.name.startsWith('table')), ...freshReference.rows])
  test(`List section ${sample.name}: menu, typing, numbering, history, reload and exports`, async ({ page }) => {
    test.setTimeout(90000);
    const run=process.env.NOFFICE_SECTION_CONTAINER_RUN || 'browser-v1';
    if (!/^browser-v\d+$/.test(run)) throw Error('Invalid container run');
    const fresh = process.env.NOFFICE_SECTION_CONTAINER_FRESH === '1';
    const root=`.local/word-section-containers/${fresh?'fresh-':''}${run}/${sample.name}`;
    await fs.mkdir(root,{recursive:true});
    const source=await fs.readFile(sample.source); expect(hash(source)).toBe(sample.sourceHash);
    const layout = process.env.NOFFICE_SECTION_CONTAINER_LAYOUT === '1';
    if (layout) await page.context().grantPermissions(['local-fonts']);
    await page.goto('/');
    await page.locator('input[type=file][multiple]').setInputFiles({name:sample.name+'.docx',buffer:source,
      mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
    const editor=page.getByRole('textbox',{name:'Document text',exact:true});
    await expect(editor).toBeVisible();
    if (fresh) {
      // Exercise the new-package writer with the imported semantic model and
      // independently authored native controls, without a retained ZIP.
      await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
      await page.evaluate(async name=>{
        const db=await new Promise<IDBDatabase>((resolve,reject)=>{const r=indexedDB.open('noffice-workspace',1);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
        try { await new Promise<void>((resolve,reject)=>{
          const tx=db.transaction('files','readwrite'),store=tx.objectStore('files'),r=store.getAll();
          r.onsuccess=()=>{const file=r.result.find((f:OfficeFile)=>f.name===name) as OfficeFile;
            delete file.original;store.put(file);};
          tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error);
        }); } finally {db.close();}
      },sample.name);
      await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();
      await page.getByRole('button',{name:sample.name,exact:true}).click();await expect(editor).toBeVisible();
    }
    const model=()=>editor.evaluate(el=>(el as HTMLElement & {editor:import('@tiptap/core').Editor}).editor.getJSON());
    const verify=async(expected:string)=>{
      const value=await editor.evaluate(el=>{
        const e=(el as HTMLElement & {editor:import('@tiptap/core').Editor}).editor;
        const live=e.state.doc.attrs.wordSectionState;
        const boundaries=new Set(live.breaks.map((b:{paragraph:number})=>b.paragraph));
        let text='',ordinal=0;
        e.state.doc.descendants(node=>{if(!node.isTextblock)return;
          text+=node.textContent+(boundaries.has(ordinal++)?'\f':'\r');return false;});
        return {text,from:e.state.selection.from,to:e.state.selection.to};
      });expect(value.text).toBe(expected);return value;
    };
    const from=await editor.evaluate((el,selection)=>{
      const e=(el as HTMLElement & {editor:import('@tiptap/core').Editor}).editor;
      const ps:number[]=[];e.state.doc.descendants((n,p)=>{if(n.isTextblock){ps.push(p);return false;}});
      const from=ps[selection.fromParagraph-1]+1+selection.fromOffset;
      e.commands.setTextSelection({from,to:ps[selection.toParagraph-1]+1+selection.toOffset});return from;
    },sample.authoredSelection);
    const before=await model();
    await page.getByRole('button',{name:'Insert',exact:true}).click();
    await page.getByRole('combobox',{name:'Insert section break',exact:true}).selectOption(sample.kind);
    const selection=await verify(sample.after);expect(selection.from).toBe(from+4);expect(selection.to).toBe(from+4);
    const after=await model();
    if (layout) {
      await waitWordLayout(editor);
      await expect(page.locator('.section-page')).toHaveCount(sample.pages);
      await clickWordText(page,editor,sample.name.includes('independent') || sample.name.includes('continuing') ? 'Delta' : 'Gamma');
      await page.keyboard.insertText('Z');
      await page.keyboard.press('Control+z');await expect.poll(model).toEqual(after);
      await editor.evaluate((el,from)=>(el as HTMLElement & {editor:import('@tiptap/core').Editor}).editor.commands.setTextSelection(from+4),from);
    }
    await editor.focus();await page.keyboard.press('Control+z');await expect.poll(model).toEqual(before);
    await page.keyboard.press('Control+y');await expect.poll(model).toEqual(after);
    await page.keyboard.insertText('X');
    await verify(sample.after.slice(0,sample.afterSelection.from)+'X'+sample.after.slice(sample.afterSelection.from));
    await page.keyboard.press('Control+z');await expect.poll(model).toEqual(after);
    const markerState=async()=>editor.locator('li').evaluateAll(items=>items.map(li=>({text:li.textContent,
      value:li.getAttribute('value'),hidden:li.getAttribute('data-word-section-list-empty')})));
    if(sample.name.startsWith('numbered-start')){
      expect(await markerState()).toEqual([
        {text:'Alpha item.',value:'1',hidden:null},{text:'',value:null,hidden:'true'},
        {text:'Beta item.',value:'2',hidden:null},{text:'Gamma item.',value:'3',hidden:null}]);
    }
    const hashes:Record<string,string>={};
    const capture=async(stage:string)=>{
      await page.getByRole('button',{name:'Export',exact:true}).click();
      const pending=page.waitForEvent('download');
      await page.getByRole('button',{name:'DOCX file Editable in Microsoft Word',exact:true}).click();
      const path=`${root}/${stage}.docx`;await(await pending).saveAs(path);hashes[stage]=hash(await fs.readFile(path));
      if (layout) {
        await waitWordLayout(editor);
        await expect(page.locator('.section-page')).toHaveCount(sample.pages);
        await page.getByRole('button',{name:'Export',exact:true}).click();
        const pdf = page.waitForEvent('download');
        await page.getByRole('button',{name:'PDF file',exact:true}).click();
        await(await pdf).saveAs(`${root}/${stage}.pdf`);
        hashes[stage+'.pdf']=hash(await fs.readFile(`${root}/${stage}.pdf`));
        if(stage==='edited') {
          await page.evaluate(()=>window.dispatchEvent(new Event('beforeprint')));
          await page.pdf({path:`${root}/print.pdf`,preferCSSPageSize:true,printBackground:true});
          await page.evaluate(()=>window.dispatchEvent(new Event('afterprint')));
          hashes['print.pdf']=hash(await fs.readFile(`${root}/print.pdf`));
          await editor.screenshot({path:`${root}/screen.png`});
          await fs.writeFile(`${root}/layout.json`,JSON.stringify(await editor.locator('p').evaluateAll(ps=>ps.map(p=>({
            text:p.textContent,marker:p.getAttribute('data-word-list-marker'),style:p.getAttribute('style'),html:p.outerHTML,
            rect:p.getBoundingClientRect().toJSON()}))),null,2));
        }
      }
    };
    await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
    await capture('edited');
    const name=await page.getByRole('textbox',{name:'File name',exact:true}).inputValue();
    await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();
    await page.getByRole('button',{name,exact:true}).click();await expect(editor).toBeVisible();
    expect(await model()).toEqual(after);await verify(sample.after);await capture('reloaded');
    await page.locator('input[type=file][multiple]').setInputFiles(root+'/edited.docx');
    await expect(editor).toBeVisible();await verify(sample.after);await capture('reimported');
    expect(hashes.reimported).toBe(hashes.edited);
    await fs.writeFile(root+'/report.json',JSON.stringify({name:sample.name,fresh,layout,sourceHash:sample.sourceHash,before,after,hashes,
      buildHash:await wordStoryBuildHash(),testHash:hash(await fs.readFile('tests/word-section-containers.spec.ts')),
      referenceHash:hash(await fs.readFile('tests/fixtures/word-section-containers/reference.json')),
      freshReferenceHash:hash(await fs.readFile('tests/fixtures/word-section-containers/fresh-reference.json')),
      scope:'Real menu, model-positioned selection, keyboard typing/history, saved model/reload, three actual DOCX stages. When layout=true: real pointer edit/undo, pagination, direct PDFs and browser print. When fresh=true: remove only original ZIP before editing to exercise new-package export. Native comparisons/return editing are independent.'},null,2));
  });
