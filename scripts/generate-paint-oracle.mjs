import fs from 'node:fs/promises';
import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
const cases = [];
for (const base of ['806040', '4472C4', 'EF3F88'])
  for (const [name, transforms] of [
    ['plain', ''],
    ['tint', '<a:tint val="40000"/>'],
    ['shade', '<a:shade val="65000"/>'],
    ['light', '<a:lumMod val="40000"/><a:lumOff val="60000"/>'],
    ['dark', '<a:lumMod val="65000"/>'],
    ['saturation', '<a:satMod val="50000"/>'],
    ['alpha', '<a:alpha val="45000"/>'],
    ['alpha-chain', '<a:alpha val="60000"/><a:alphaMod val="50000"/><a:alphaOff val="10000"/>'],
  ])
    cases.push({
      name: `${base}-${name}`,
      xml: `<a:srgbClr val="${base}">${transforms}</a:srgbClr>`,
    });
const pptx = new PptxGenJS();
for (const c of cases)
  pptx
    .addSlide()
    .addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 3, h: 2, fill: { color: '123456' } });
const zip = await JSZip.loadAsync(await pptx.write({ outputType: 'arraybuffer' }));
for (let i = 0; i < cases.length; i++) {
  const path = `ppt/slides/slide${i + 1}.xml`;
  zip.file(
    path,
    (await zip.file(path).async('string')).replace(
      /<a:solidFill>[\s\S]*?<\/a:solidFill>/,
      `<a:solidFill>${cases[i].xml}</a:solidFill>`,
    ),
  );
}
await fs.mkdir('.local/pptx-validation', { recursive: true });
await fs.writeFile(
  '.local/pptx-validation/paint-oracle.pptx',
  await zip.generateAsync({ type: 'nodebuffer' }),
);
await fs.writeFile('tests/fixtures/paint-oracle-cases.json', JSON.stringify(cases, null, 2) + '\n');
