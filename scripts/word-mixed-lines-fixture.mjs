import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';

// Authored test content only. Native Word saves the checked fixture separately.
const zip = await JSZip.loadAsync(
  await Packer.toBuffer(
    new Document({
      sections: [{ children: [new Paragraph('Fixture')] }],
    }),
  ),
);
const font = (size) => `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${size * 2}"/>`;
const run = (text, size) =>
  `<w:r><w:rPr>${font(size)}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const cases = [
  ['auto', 240, 10],
  ['auto', 360, 10],
  ['auto', 240, 40],
  ['atLeast', 360, 10],
  ['atLeast', 1000, 10],
  ['atLeast', 360, 40],
];
const body = cases
  .map(
    ([rule, line, mark]) =>
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="${rule}"/><w:widowControl w:val="0"/><w:rPr>${font(mark)}</w:rPr></w:pPr>${run('Small ', 10)}${run('Large', 30)}${run(' end', 10)}</w:p>`,
  )
  .join('');
zip.file(
  'word/document.xml',
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr>${font(10)}</w:rPr></w:pPr>${run('Sentinel', 10)}</w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
);
await fs.mkdir('.local/word-mixed-lines', { recursive: true });
await fs.writeFile(
  '.local/word-mixed-lines/authored.docx',
  await zip.generateAsync({ type: 'nodebuffer' }),
);
