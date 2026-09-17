import { expect, it } from 'vitest';
import { authoredWordNumbering } from './docx-authored-numbering';
import type { WordListDefinition, WordNumbering } from './word-list-layout';

const definition: WordListDefinition = { numId:'7',level:0,start:4,format:'decimal',text:'%1.',
  left:720,hanging:360,font:'Arial',size:12 };
const metadata = (...definitions: WordListDefinition[]): WordNumbering => ({version:1,
  paragraphs:definitions.map((definition,i)=>({source:`0:${i}`,numbering:{status:'resolved',definition}}))});
const body = (html: string) => new DOMParser().parseFromString(html,'text/html').body;
const list = (index: number) => `<ol><li><p data-source-paragraph="0:${index}">Item</p></li></ol>`;

it('keeps a source counter across wrappers and independent starts across numIds', () => {
  const dom=body(list(0)+list(1)+list(2));
  const plan=authoredWordNumbering([dom],metadata(definition,definition,{...definition,numId:'8',start:9}));
  const ps=[...dom.querySelectorAll('p')];
  expect(plan.reference(ps[0])).toBe(plan.reference(ps[1]));
  expect(plan.reference(ps[2])).not.toBe(plan.reference(ps[0]));
  expect(plan.config.map(c=>c.levels[0].start)).toEqual([4,9]);
  expect(plan.config[0].levels[0].style).toEqual({paragraph:{indent:{left:720,hanging:360}},run:{font:'Arial',size:24,sizeComplexScript:false}});
  expect(plan.reference(ps[0].cloneNode(true) as Element)).toBe(plan.reference(ps[0]));
});

it('assigns independent authored lists and preserves starts before section cloning', () => {
  const dom=body('<ol start="4"><li><p>First</p></li></ol><p>Gap</p><ol><li><p>Second</p></li></ol>');
  const plan=authoredWordNumbering([dom]);
  expect(plan.config.map(c=>c.levels[0].start)).toEqual([4,1]);
  expect(new Set([...dom.querySelectorAll('li p')].map(p=>plan.reference(p))).size).toBe(2);
  expect(plan.reference(dom.children[1])).toBeUndefined();
});

it('never trusts export-only reference attributes from input', () => {
  const dom=body('<p data-noffice-export-numbering="source-list-7">Plain</p>'+list(0));
  const plan=authoredWordNumbering([dom],metadata(definition));
  expect(plan.reference(dom.children[0])).toBeUndefined();
  expect(plan.reference(dom.querySelector('li p')!)).toBe('source-list-7');
});

it('writes literal qualified bullets instead of the library bullet font', () => {
  const dom=body(list(0).replaceAll('ol>','ul>'));
  const plan=authoredWordNumbering([dom],metadata({...definition,format:'bullet',text:'\u2022'}));
  expect(plan.config[0].levels[0]).toMatchObject({format:'bullet',text:'\u2022',style:{run:{font:'Arial',size:24}}});
});

it('rejects unresolved, missing, conflicting and nested source identities', () => {
  const unknown: WordNumbering={version:1,paragraphs:[{source:'0:0',numbering:{status:'unsupported',reason:'Custom marker'}}]};
  expect(()=>authoredWordNumbering([body(list(0))],unknown)).toThrow(/retained DOCX/);
  expect(()=>authoredWordNumbering([body(list(0))])).toThrow(/retained DOCX/);
  expect(()=>authoredWordNumbering([body('<ol><li><p data-source-paragraph="0:0">A</p></li><li><p>B</p></li></ol>')],metadata(definition))).toThrow(/retained DOCX/);
  expect(()=>authoredWordNumbering([body(list(0)+list(1))],metadata(definition,{...definition,start:2}))).toThrow(/retained DOCX/);
  expect(()=>authoredWordNumbering([body('<ol><li><p>Outer</p>'+list(0)+'</li></ol>')],metadata(definition))).toThrow(/retained DOCX/);
  expect(()=>authoredWordNumbering([body(list(0).replaceAll('ol>','ul>'))],metadata(definition))).toThrow(/retained DOCX/);
});

it('rejects invalid authored starts and unsupported numbering instead of silently coercing them', () => {
  for (const attrs of ['start="-1"','start="1000001"','start="1.5"','reversed','type="I"'])
    expect(()=>authoredWordNumbering([body(`<ol ${attrs}><li><p>A</p></li></ol>`)]),attrs).toThrow(/retained DOCX/);
});

it('rejects marker sizes that the OOXML writer would truncate to a half point', () => {
  expect(()=>authoredWordNumbering([body(list(0))],metadata({...definition,size:12.25}))).toThrow();
  expect(authoredWordNumbering([body(list(0))],metadata({...definition,size:12.5})).config[0].levels[0].style?.run?.size).toBe(25);
});
