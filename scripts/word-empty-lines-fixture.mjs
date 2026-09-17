import fs from 'node:fs/promises';
import JSZip from 'jszip';
const root = '.local/word-empty-lines';
await fs.mkdir(root, { recursive: true });
if (process.argv.includes('--prepare')) {
  const zip = await JSZip.loadAsync(await fs.readFile(`${root}/source.docx`));
  const core = await zip.file('docProps/core.xml').async('string');
  zip.file(
    'docProps/core.xml',
    core.replace(
      /(<(?:dc:creator|cp:lastModifiedBy)>).*?(<\/(?:dc:creator|cp:lastModifiedBy)>)/g,
      '$1Noffice$2',
    ),
  );
  await fs.writeFile(`${root}/source.docx`, await zip.generateAsync({ type: 'nodebuffer' }));
  process.exit(0);
}
const zip = await JSZip.loadAsync(await fs.readFile('tests/fixtures/word-baselines.docx'));
const font = (size) => `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${size * 2}"/>`;
const paragraph = (text, mark = 10) =>
  `<w:p><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/><w:rPr>${font(mark)}</w:rPr></w:pPr>${text
    .split('\n')
    .map((t) => `<w:r><w:rPr>${font(10)}</w:rPr><w:t>${t}</w:t></w:r>`)
    .join(`<w:r><w:rPr>${font(10)}</w:rPr><w:br/></w:r>`)}</w:p>`;
const cases = [
  ['\nAlpha', 10],
  ['Alpha\n', 10],
  ['Alpha\n\nBeta', 10],
  ['\n\n', 10],
  ['Alpha\nBeta', 10],
  ['\nAlpha', 40],
  ['Alpha\n', 40],
  ['Alpha\n\nBeta', 40],
  ['\n\n', 40],
];
zip.file(
  'word/document.xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${cases.map(([text, mark], i) => paragraph(text, mark) + paragraph(`Sentinel${i + 1}`)).join('')}<w:sectPr><w:pgSz w:w="9000" w:h="31680"/><w:pgMar w:top="900" w:right="900" w:bottom="900" w:left="900" w:header="400" w:footer="400" w:gutter="0"/></w:sectPr></w:body></w:document>`,
);
await fs.writeFile(`${root}/authored.docx`, await zip.generateAsync({ type: 'nodebuffer' }));
