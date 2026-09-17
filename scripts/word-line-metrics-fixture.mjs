import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';

export async function lineMetricsFixture() {
  return fs.readFile(new URL('../tests/fixtures/word-line-metrics.docx', import.meta.url));
}

/** Authored variant: paragraph ancestry, character/direct overrides and distinct marks. */
export async function fontCascadeFixture() {
  return fs.readFile(new URL('../tests/fixtures/word-font-cascade.docx', import.meta.url));
}

export async function rawFontCascadeFixture() {
  const zip = await JSZip.loadAsync(await lineMetricsFixture());
  const stylesPart = zip.file('word/styles.xml'),
    mainPart = zip.file('word/document.xml');
  if (!stylesPart || !mainPart) throw Error('Incomplete baseline font fixture');
  let styles = await stylesPart.async('string');
  styles = styles.replace(
    '</w:styles>',
    '<w:style w:type="paragraph" w:styleId="NofficeBase"><w:rPr><w:sz w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="NofficeDerived"><w:basedOn w:val="NofficeBase"/></w:style><w:style w:type="character" w:styleId="NofficeAccent"><w:rPr><w:sz w:val="36"/></w:rPr></w:style></w:styles>',
  );
  let index = 0;
  const xml = (await mainPart.async('string')).replace(/<w:p(?=[ >])[\s\S]*?<\/w:p>/g, (p) => {
    const i = index++;
    if (i === 1)
      return p
        .replace('<w:pPr>', '<w:pPr><w:pStyle w:val="NofficeDerived"/>')
        .replace('</w:pPr>', '<w:rPr><w:sz w:val="48"/></w:rPr></w:pPr>')
        .replace(
          '<w:r><w:t>Small again',
          '<w:r><w:rPr><w:rStyle w:val="NofficeAccent"/></w:rPr><w:t>Small again',
        );
    if (i === 3) return p.replace('</w:pPr>', '<w:rPr><w:sz w:val="60"/></w:rPr></w:pPr>');
    return p;
  });
  if (index !== 6) throw Error('Unexpected font fixture paragraph count');
  zip.file('word/styles.xml', styles);
  zip.file('word/document.xml', xml);
  return zip.generateAsync({ type: 'nodebuffer' });
}

export async function rawLineMetricsFixture() {
  const zip = await JSZip.loadAsync(
    await Packer.toBuffer(
      new Document({
        styles: { default: { document: { run: { font: 'Arial', size: 20 } } } },
        sections: [{ children: [new Paragraph('Fixture')] }],
      }),
    ),
  );
  /** @param {string} text @param {number} [size] */
  const run = (text, size = 20) =>
    `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${size}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
  /** @param {string} rule @param {number} amount @param {string} content */
  const p = (rule, amount, content) =>
    `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${amount}" w:lineRule="${rule}"/><w:widowControl w:val="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:pPr>${content}</w:p>`;
  const mixed =
    run('Small line') +
    '<w:r><w:br/></w:r>' +
    run('Large line', 40) +
    '<w:r><w:br/></w:r>' +
    run('Small again') +
    '<w:r><w:br/></w:r>';
  const section =
    '<w:pgSz w:w="5000" w:h="4000"/><w:pgMar w:top="400" w:bottom="400" w:left="400" w:right="400" w:header="200" w:footer="200" w:gutter="0"/>';
  let body =
    p('exact', 240, mixed) +
    p('atLeast', 240, mixed) +
    p('auto', 240, mixed) +
    p('atLeast', 360, '') +
    p('auto', 360, run(Array.from({ length: 35 }, (_, i) => `word${i + 1}`).join(' ')));
  // End the first section on the final paragraph; keep source paragraph identities.
  const at = body.lastIndexOf('</w:pPr>');
  body = body.slice(0, at) + `<w:sectPr>${section}</w:sectPr>` + body.slice(at);
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}${p('exact', 240, run('Final section.'))}<w:sectPr><w:type w:val="nextPage"/>${section}</w:sectPr></w:body></w:document>`,
  );
  return zip.generateAsync({ type: 'nodebuffer' });
}
