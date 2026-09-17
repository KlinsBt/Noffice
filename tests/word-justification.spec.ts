import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import arial from './fixtures/native-word-justification-browser.json' with { type: 'json' };
import calibri from './fixtures/native-word-justification-calibri-browser.json' with { type: 'json' };
import { wordStoryBuildHash } from './word-story-artifacts';

const profile = process.env.NOFFICE_JUSTIFICATION_PROFILE || 'arial-ten';
if (!['arial-ten', 'calibri-eleven'].includes(profile)) throw Error('Invalid justification profile');
const reference = profile === 'calibri-eleven' ? calibri : arial;
for (const row of reference.rows) test(`Word modern justification ${row.name}`, async ({ page }) => {
  const run = process.env.NOFFICE_JUSTIFICATION_RUN || 'browser-v1';
  if (!/^browser-v\d+$/.test(run)) throw Error('Invalid justification run');
  const root = `.local/word-story-paragraph-layout/justification-${profile}/${run}/${row.name}`;
  await fs.mkdir(root, { recursive: true });
  const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');
  expect(hash(await fs.readFile(row.fixture))).toBe(row.sourceHash);
  const errors: string[] = []; page.on('pageerror', (error) => errors.push(error.message));
  await page.context().grantPermissions(['local-fonts']); await page.goto('/');
  const name = 'Justification ' + row.name;
  await page.locator('input[type=file][multiple]').setInputFiles({ name:name+'.docx',buffer:await fs.readFile(row.fixture),mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  await expect(page.locator('.section-page')).toHaveCount(row.pages);
  const target = row.kind === 'body'
    ? page.locator('.tiptap.document-page p').filter({ hasText: row.text }).first()
    : page.locator('.word-story-measure-host p').filter({ hasText: row.text }).first();
  await expect(target).toHaveAttribute('data-word-justification', 'modern');
  const read = () => target.evaluate((paragraph) => {
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    const lines: { y: number; end: number; text: string }[] = [];
    let offset = 0;
    while (walker.nextNode()) {
      const node = walker.currentNode;
      if (node.parentElement?.closest('.ProseMirror-widget')) continue;
      for (let i = 0; i < (node.textContent?.length || 0); i++,offset++) {
        const text = node.textContent![i]; if (/\s/.test(text)) continue;
        const range = document.createRange(); range.setStart(node,i); range.setEnd(node,i+1);
        const y = range.getBoundingClientRect().top;
        if (!lines.length || Math.abs(lines.at(-1)!.y-y)>10) lines.push({ y,end:offset+1,text });
        else { lines.at(-1)!.end=offset+1; lines.at(-1)!.text+=text; }
      }
    }
    return { lines,text:paragraph.textContent,html:paragraph.innerHTML };
  });
  await expect.poll(async () => (await read()).lines.map((line) => line.end)).toEqual(row.ends);
  const rendered = await read(); expect(rendered.text).toBe(row.text);
  const outputs: Record<string,string> = {};
  const download = async (label: string, file: string) => {
    await page.getByRole('button', { name:'Export',exact:true }).click();
    const waiting=page.waitForEvent('download');await page.getByRole('button',{name:label,exact:true}).click();
    await (await waiting).saveAs(root+'/'+file);
    outputs[file]=hash(await fs.readFile(root+'/'+file));
  };
  for (const [label,file] of [['DOCX file Editable in Microsoft Word','original.docx'],['PDF file','actual.pdf']]) {
    await download(label,file);
  }
  expect(hash(await fs.readFile(root+'/original.docx'))).toBe(row.sourceHash);
  let editor = page.getByRole('textbox',{name:'Document text',exact:true});
  if (row.kind !== 'body') {
    await page.getByRole('button',{name:'Insert',exact:true}).click();
    await page.getByRole('button',{name:'Headers and footers',exact:true}).click();
    await page.getByRole('button',{name:new RegExp('^Section 1 \\u2014 Default '+row.kind)}).click();
    editor=page.getByRole('textbox',{name:'Header or footer text',exact:true});
  }
  const editedParagraph=editor.locator('p').filter({hasText:row.text}).first();
  await editor.focus();await page.keyboard.press('Control+Home');
  if(row.kind!=='body') { await page.keyboard.press('Control+ArrowDown');await page.keyboard.press('Control+ArrowDown');await page.keyboard.press('Home'); }
  await page.keyboard.type('New ');await expect(editedParagraph).toHaveText('New '+row.text);
  const toolbar=row.kind==='body'?page:page.getByRole('dialog');
  await toolbar.getByRole('button',{name:'Undo',exact:true}).click();await expect(editedParagraph).toHaveText(row.text);
  await expect(toolbar.getByRole('button',{name:'Undo',exact:true})).toBeDisabled();
  await toolbar.getByRole('button',{name:'Redo',exact:true}).click();await expect(editedParagraph).toHaveText('New '+row.text);
  if(row.kind!=='body') await page.getByRole('button',{name:'Apply header or footer',exact:true}).click();
  await expect(target).toHaveText('New '+row.text);
  await expect(target).toHaveAttribute('data-word-justification','modern');
  await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  const edited=await read();
  await download('DOCX file Editable in Microsoft Word','edited.docx');
  await download('PDF file','edited.pdf');
  await expect(page.getByText('Saved on this device',{exact:true})).toBeVisible();
  await page.reload();await page.getByRole('button',{name:'Recent files',exact:true}).click();
  await page.getByRole('button',{name,exact:true}).click();
  await expect(page.locator('.section-page')).toHaveCount(row.pages);
  await expect(target).toHaveText('New '+row.text);
  await expect.poll(async () => (await read()).lines.map((line) => line.end)).toEqual(edited.lines.map((line)=>line.end));
  await download('DOCX file Editable in Microsoft Word','reloaded.docx');
  await download('PDF file','reloaded.pdf');
  expect(errors).toEqual([]);
  await fs.writeFile(root+'/browser-report.json',JSON.stringify({name:row.name,sourceHash:row.sourceHash,
    nativePdfHash:row.nativePdfHash,pdfHash:hash(await fs.readFile(root+'/actual.pdf')),outputs,rendered,edited,actions:{prefix:'New '},errors,
    buildHash:await wordStoryBuildHash(),testHash:hash(await fs.readFile('tests/word-justification.spec.ts'))},null,2));
});
