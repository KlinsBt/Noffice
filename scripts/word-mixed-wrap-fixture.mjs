import fs from 'node:fs/promises';
import JSZip from 'jszip';

const root = '.local/word-mixed-wrap';
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
// MIT-authored synthetic text, using the existing Noffice package shell.
const zip = await JSZip.loadAsync(await fs.readFile('tests/fixtures/word-baselines.docx'));
const font = (size) => `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${size * 2}"/>`;
const run = (text, size = 10) =>
  `<w:r><w:rPr>${font(size)}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (text, size = 10, line = 360) =>
  `<w:p><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="auto"/><w:rPr>${font(size)}</w:rPr></w:pPr>${text}</w:p>`;
const words = (prefix) =>
  Array.from({ length: 20 }, (_, i) =>
    run(`${prefix}${i + 1}${i === 19 ? '' : ' '}`, [10, 10, 20, 10, 30, 10, 10][i % 7]),
  ).join('');
const body =
  paragraph(run('Start'), 10, 240) +
  paragraph(words('alpha')) +
  paragraph(words('beta'), 40) +
  paragraph(run('End'), 10, 240);
zip.file(
  'word/document.xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="6000" w:h="22000"/><w:pgMar w:top="900" w:right="900" w:bottom="900" w:left="900" w:header="400" w:footer="400" w:gutter="0"/></w:sectPr></w:body></w:document>`,
);
await fs.writeFile(`${root}/authored.docx`, await zip.generateAsync({ type: 'nodebuffer' }));
