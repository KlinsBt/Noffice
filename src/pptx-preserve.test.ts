import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { newFile, textElement } from './model';
import { contentFingerprint } from './office-preservation';
import { exportAddedPresentationObjects } from './pptx-preserve';

async function setup() {
  const file = newFile('powerpoint');
  if (file.content.kind !== 'powerpoint') throw Error();
  file.content.slides = [
    {
      id: 'slide',
      sourcePath: 'ppt/slides/slide1.xml',
      background: '#fff',
      notes: '',
      elements: [textElement('Original title', { id: 'title', sourceShapeId: '2' })],
    },
  ];
  const zip = new JSZip();
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="9144000" cy="6858000"/></p:presentation>',
  );
  zip.file(
    'ppt/slides/slide1.xml',
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Original"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Original title</a:t></a:r></a:p></p:txBody></p:sp><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Chart"/></p:nvGraphicFramePr></p:graphicFrame></p:spTree></p:cSld><p:timing/></p:sld>',
  );
  zip.file('ppt/charts/chart1.xml', '<retained-chart/>');
  file.original = {
    name: 'source.pptx',
    data: await zip.generateAsync({ type: 'arraybuffer' }),
    contentFingerprint: await contentFingerprint(file.content),
  };
  return file;
}
async function bytes(blob: Blob) {
  return new Promise<ArrayBuffer>((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.readAsArrayBuffer(blob);
  });
}
describe('adding objects to retained presentations', () => {
  it('adds escaped text with unique IDs while retaining chart parts, slide timing and dimensions', async () => {
    const file = await setup();
    if (file.content.kind !== 'powerpoint') throw Error();
    file.content.slides[0].elements.push(textElement('A < B & C', { bold: true }));
    const result = await exportAddedPresentationObjects(file);
    expect(result).not.toBeNull();
    const zip = await JSZip.loadAsync(await bytes(result!));
    expect(await zip.file('ppt/charts/chart1.xml')!.async('string')).toBe('<retained-chart/>');
    expect(await zip.file('ppt/presentation.xml')!.async('string')).toContain('cx="9144000"');
    const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
    expect(slide).toContain('A &lt; B &amp; C');
    expect(slide).toContain('id="4"');
    expect(slide).toContain('p:timing');
    expect(slide).toContain('Original title');
  });
  it('declines this limited path when existing text changed', async () => {
    const file = await setup();
    if (file.content.kind !== 'powerpoint') throw Error();
    file.content.slides[0].elements[0].text = 'Changed existing content';
    file.content.slides[0].elements.push(textElement('Addition'));
    expect(await exportAddedPresentationObjects(file)).toBeNull();
  });
});
