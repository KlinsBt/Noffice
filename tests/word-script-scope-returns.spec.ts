import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import reference from './fixtures/native-word-script-scopes.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

for (const source of reference.rows) test(`Word script scopes ${source.name}: native edit returns through typing, history and reload`, async ({ page }) => {
  test.skip(process.env.NOFFICE_SCRIPT_SCOPE_RETURN !== '1', 'Requires independently edited installed Word output');
  const sourceRun = process.env.NOFFICE_SCRIPT_SCOPE_SOURCE_RUN || 'browser-v6';
  const run = process.env.NOFFICE_SCRIPT_SCOPE_RETURN_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(sourceRun) || !/^browser-v\d+$/.test(run)) throw Error('Invalid evidence folder');
  const nativeRoot = `.local/word-script-scopes/${sourceRun}/${source.name}/native-v1`;
  const root = nativeRoot + '/' + run;await fs.mkdir(root, {recursive:true});
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  const nativeBytes = await fs.readFile(nativeRoot+'/native-report.json');
  const native = JSON.parse(nativeBytes.toString().replace(/^\uFEFF/,''));
  const expected = native.stages.find((stage: {stage:string}) => stage.stage === 'returned');
  const input = await fs.readFile(nativeRoot+'/returned.docx');expect(hash(input)).toBe(expected.docxHash);
  expect(hash(await fs.readFile(nativeRoot+'/returned.pdf'))).toBe(expected.pdfHash);
  expect(native.executableHash).toBe(reference.executableHash);
  const errors: string[] = [];page.on('pageerror', error=>errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']);await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({name:`Script scopes return ${source.name}.docx`,buffer:input,
    mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const body = source.kind === 'Body',kind=source.kind==='Headers'?'header':'footer';
  const openStory = async () => {
    await page.getByRole('button',{name:'Insert',exact:true}).click();
    await page.getByRole('button',{name:'Headers and footers',exact:true}).click();
    await page.getByRole('button',{name:new RegExp(`^Section 1 \\u2014 Default ${kind}`)}).click();
  };
  if (!body) await openStory();
  const editor=page.getByRole('textbox',{name:body?'Document text':'Header or footer text',exact:true});
  const characters=()=>editor.evaluate(element=>{
    type Node={type:string;text?:string;attrs:Record<string,unknown>;content?:Node[];marks?:{type:string;attrs:Record<string,unknown>}[]};
    const p=(element as HTMLElement&{editor:{getJSON():Node}}).editor.getJSON().content![0];
    const read=(text:string,marks:Node['marks']=[],paragraph=false)=>{
      const attrs=marks?.find(m=>m.type==='textStyle')?.attrs;
      return {text,family:String(attrs?.fontFamily||p.attrs.paragraphFontFamily).replace(/^["']|["']$/g,''),size:parseFloat(String(attrs?.fontSize||p.attrs.paragraphFontSize)),
        superscript:(paragraph?p.attrs.paragraphScript==='superscript':marks?.some(m=>m.type==='superscript'))?-1:0,subscript:(paragraph?p.attrs.paragraphScript==='subscript':marks?.some(m=>m.type==='subscript'))?-1:0};
    };
    return [...(p.content||[]).flatMap(n=>[...(n.type==='wordTab'?'\t':n.text||'')].map(c=>read(c,n.marks))),read('\r',[],true)];
  });
  await expect.poll(characters).toEqual(expected.characters);
  const originalText=expected.characters.map((c:{text:string})=>c.text).join('').replace(/\r$/,'');
  await editor.focus();await page.keyboard.press('Control+Home');await page.keyboard.insertText('Z');
  await expect(editor.locator('p').first()).toHaveText('Z'+originalText);
  await page.keyboard.press('Control+z');await expect.poll(characters).toEqual(expected.characters);
  await page.keyboard.press('Control+y');await expect(editor.locator('p').first()).toHaveText('Z'+originalText);
  await page.keyboard.press('Control+z');await expect.poll(characters).toEqual(expected.characters);
  if (!body) await page.getByRole('button',{name:'Apply header or footer',exact:true}).click();
  const outputs:Record<string,string>={};
  const download=async(name:string,label:string)=>{
    await page.getByRole('button',{name:'Export',exact:true}).click();const pending=page.waitForEvent('download');
    await page.getByRole('button',{name:label,exact:true}).click();await(await pending).saveAs(root+'/'+name);
    outputs[name]=hash(await fs.readFile(root+'/'+name));
  };
  await download('returned.docx','DOCX file Editable in Microsoft Word');await download('returned.pdf','PDF file');
  expect(await fs.readFile(root+'/returned.docx')).toEqual(input);
  await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();
  await page.getByRole('button',{name:`Script scopes return ${source.name}`,exact:true}).click();
  if (!body) await openStory();await expect.poll(characters).toEqual(expected.characters);
  if (!body) await page.getByRole('button',{name:'Apply header or footer',exact:true}).click();
  await download('reloaded.pdf','PDF file');
  await page.goto('/');await page.locator('input[type=file][multiple]').setInputFiles(root+'/returned.docx');
  if (!body) await openStory();await expect.poll(characters).toEqual(expected.characters);expect(errors).toEqual([]);
  await fs.writeFile(root+'/browser-report.json',JSON.stringify({outputs,errors,sourceHash:hash(input),nativeReceiptHash:hash(nativeBytes),
    buildHash:await wordStoryBuildHash(),testHash:hash(await fs.readFile('tests/word-script-scope-returns.spec.ts')),
    referenceHash:hash(await fs.readFile('tests/fixtures/native-word-script-scopes.json'))},null,2));
});
