import fs from 'node:fs/promises';
import JSZip from 'jszip';

if (process.argv.includes('--prepare')) {
  const path = '.local/word-wrapped-leading/source.docx';
  const saved = await JSZip.loadAsync(await fs.readFile(path));
  const core = await saved.file('docProps/core.xml').async('string');
  saved.file(
    'docProps/core.xml',
    core.replace(
      /(<(?:dc:creator|cp:lastModifiedBy)>).*?(<\/(?:dc:creator|cp:lastModifiedBy)>)/g,
      '$1Noffice$2',
    ),
  );
  await fs.writeFile(path, await saved.generateAsync({ type: 'nodebuffer' }));
  process.exit(0);
}

// Noffice-authored words from the existing line-metrics fixture; no user content.
const zip = await JSZip.loadAsync(await fs.readFile('tests/fixtures/word-baselines.docx'));
const font = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/>';
const run = (text) => `<w:r><w:rPr>${font}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (contents, multiple = 360) =>
  `<w:p><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="${multiple}" w:lineRule="auto"/><w:rPr>${font}</w:rPr></w:pPr>${contents}</w:p>`;
const words = Array.from({ length: 35 }, (_, i) => `word${i + 1}`).join(' ');
const br = `<w:r><w:rPr>${font}</w:rPr><w:br/></w:r>`;
const body =
  paragraph(run('Start'), 240) +
  paragraph(run(words)) +
  paragraph(['First line', 'Second line', 'Third line'].map(run).join(br)) +
  paragraph(run('End'), 240);
zip.file(
  'word/document.xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="6000" w:h="16840"/><w:pgMar w:top="900" w:right="900" w:bottom="900" w:left="900" w:header="400" w:footer="400" w:gutter="0"/></w:sectPr></w:body></w:document>`,
);
await fs.mkdir('.local/word-wrapped-leading', { recursive: true });
await fs.writeFile(
  '.local/word-wrapped-leading/authored.docx',
  await zip.generateAsync({ type: 'nodebuffer' }),
);
