import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';

/** @param {{ firstLines?: number, terminalBreak?: boolean, keepSecond?: boolean }} [options] */
export async function paginationFixture({
  firstLines = 20,
  terminalBreak = false,
  keepSecond = true,
} = {}) {
  if (!Number.isInteger(firstLines) || firstLines < 1 || firstLines > 100)
    throw new Error('Invalid fixture line count');
  const zip = await JSZip.loadAsync(
    await Packer.toBuffer(
      new Document({
        styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
        sections: [{ children: [new Paragraph('Fixture')] }],
      }),
    ),
  );
  const spacing =
    '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/><w:widowControl w:val="0"/>';
  /** @param {string} text @param {boolean} [bold] */
  const run = (text, bold = false) =>
    `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  /** @param {string} prefix @param {number} count */
  const lines = (prefix, count) =>
    Array.from(
      { length: count },
      (_, i) =>
        (i ? '<w:r><w:br/></w:r>' : '') +
        run(`${prefix} ${String(i + 1).padStart(2, '0')} `) +
        run('text', i % 3 === 0),
    ).join('');
  /** @param {number} width */
  const section = (width) =>
    `<w:pgSz w:w="${width}" w:h="4000"/><w:pgMar w:top="400" w:bottom="400" w:left="400" w:right="400" w:header="200" w:footer="200" w:gutter="0"/>`;
  const wrapped = Array.from(
    { length: 110 },
    (_, i) => `word${String(i + 1).padStart(3, '0')}`,
  ).join(' ');
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
    <w:p><w:pPr>${spacing}</w:pPr>${lines('Line', firstLines)}${terminalBreak ? '<w:r><w:br/></w:r>' : ''}</w:p>
    <w:p><w:pPr><w:keepLines w:val="${keepSecond ? 1 : 0}"/>${spacing}</w:pPr>${lines('Keep', 8)}</w:p>
    <w:p><w:pPr>${spacing}<w:sectPr>${section(5000)}</w:sectPr></w:pPr>${run(wrapped)}</w:p>
    <w:p><w:pPr>${spacing}</w:pPr>${run('Final section.')}</w:p><w:sectPr><w:type w:val="nextPage"/>${section(6000)}</w:sectPr>
    </w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'nodebuffer' });
}
