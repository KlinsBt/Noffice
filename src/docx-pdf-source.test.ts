import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { assertWordPdfSourceDecorations } from './docx-pdf-source';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function source(run: string, options: { styles?: string; paragraph?: string; story?: string } = {}) {
  const zip = new JSZip();
  zip.file('word/document.xml', `<w:document xmlns:w="${ns}"><w:body><w:p><w:pPr>${options.paragraph || ''}</w:pPr><w:r><w:rPr>${run}</w:rPr><w:t>A</w:t></w:r></w:p></w:body></w:document>`);
  if (options.styles) zip.file('word/styles.xml', `<w:styles xmlns:w="${ns}">${options.styles}</w:styles>`);
  if (options.story) zip.file('word/custom-story.xml', `<w:hdr xmlns:w="${ns}"><w:p><w:r><w:rPr>${options.story}</w:rPr><w:t>Header</w:t></w:r></w:p></w:hdr>`);
  return zip.generateAsync({ type: 'arraybuffer' });
}
describe('retained DOCX PDF decoration loss boundary', () => {
  it.each(['', '<w:u/>', '<w:u w:val="single" w:color="auto"/>', '<w:strike/>',
    '<w:u w:val="none" w:color="FF0000"/><w:dstrike w:val="0"/>'])('accepts supported effective marks %s without changing originals', async run => {
    const bytes = await source(run), before = bytes.slice(0);
    await expect(assertWordPdfSourceDecorations(bytes)).resolves.toBeUndefined();
    expect(bytes).toEqual(before);
  });
  it.each(['<w:u w:val="double"/>', '<w:u w:val="wave"/>', '<w:u w:val="words"/>',
    '<w:u w:val="single" w:color="FF0000"/>', '<w:u w:val="single" w:themeColor="accent1"/>',
    '<w:dstrike/>'])('rejects retained source loss %s', async run => {
    await expect(assertWordPdfSourceDecorations(await source(run))).rejects.toThrow(/complex underline or double-strikethrough/);
  });
  it('resolves used style chains, respects a direct none override, and ignores unused styles', async () => {
    const styles = '<w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:u w:val="double"/></w:rPr></w:style>'
      + '<w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style>';
    await expect(assertWordPdfSourceDecorations(await source('', { styles }))).resolves.toBeUndefined();
    const options = { styles, paragraph: '<w:pStyle w:val="Child"/>' };
    await expect(assertWordPdfSourceDecorations(await source('', options))).rejects.toThrow(/complex underline/);
    await expect(assertWordPdfSourceDecorations(await source('<w:u w:val="none"/>', options))).resolves.toBeUndefined();
    await expect(assertWordPdfSourceDecorations(await source('', { story: '<w:dstrike/>' }))).rejects.toThrow(/double-strikethrough/);
  });
  it('rejects malformed archives and XML before processing typography', async () => {
    await expect(assertWordPdfSourceDecorations(new ArrayBuffer(12))).rejects.toThrow();
    const zip = new JSZip(); zip.file('word/document.xml', '<!DOCTYPE x><x/>');
    await expect(assertWordPdfSourceDecorations(await zip.generateAsync({ type: 'arraybuffer' }))).rejects.toThrow(/DTD/);
  });
});
