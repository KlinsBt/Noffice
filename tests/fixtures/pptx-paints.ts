import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
export async function paintPresentationFixture(): Promise<ArrayBuffer> {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  const slide = pptx.addSlide();
  slide.addText('Filled text', {
    objectName: 'Filled text',
    x: 1,
    y: 1,
    w: 4,
    h: 2,
    fontSize: 24,
    fill: { color: 'DDEEAA', transparency: 50 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    objectName: 'No fill',
    x: 1,
    y: 4,
    w: 2,
    h: 1,
    fill: { color: 'FFFFFF' },
    line: { color: 'FF0000', width: 2 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    objectName: 'Theme fill',
    x: 7,
    y: 1,
    w: 3,
    h: 2,
    fill: { color: '123456' },
  });
  slide.addNotes('Paint retains notes');
  pptx.addSlide().addText('Untouched slide', { x: 1, y: 1, w: 5, h: 1 });
  const zip = await JSZip.loadAsync(
    (await pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer,
  );
  const path = 'ppt/slides/slide1.xml';
  const parts = (await zip.file(path)!.async('string')).split('</p:sp>');
  for (let i = 0; i < parts.length; i++) {
    if (parts[i].includes('name="No fill"'))
      parts[i] = parts[i].replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, '<a:noFill/>');
    if (parts[i].includes('name="Theme fill"'))
      parts[i] = parts[i].replace(
        /<a:solidFill>[\s\S]*?<\/a:solidFill>/,
        '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>',
      );
  }
  zip.file(path, parts.join('</p:sp>'));
  for (const entry of Object.keys(zip.files))
    if (/^ppt\/(slides|slideLayouts)\/[^/]+\.xml$/.test(entry))
      zip.file(
        entry,
        (await zip.file(entry)!.async('string')).replace(/<p:bg\b[^>]*>[\s\S]*?<\/p:bg>/g, ''),
      );
  const masterPath = 'ppt/slideMasters/slideMaster1.xml';
  const master = await zip.file(masterPath)!.async('string');
  zip.file(
    masterPath,
    master
      .replace(/<p:bg>[\s\S]*?<\/p:bg>/, '')
      .replace(
        /(<p:cSld\b[^>]*>)/,
        '$1<p:bg><p:bgPr><a:solidFill><a:srgbClr val="BBDDEE"/></a:solidFill></p:bgPr></p:bg>',
      ),
  );
  const themePath = 'ppt/theme/theme1.xml';
  zip.file(
    themePath,
    (await zip.file(themePath)!.async('string')).replace(
      /<a:accent1>[\s\S]*?<\/a:accent1>/,
      '<a:accent1><a:srgbClr val="345678"/></a:accent1>',
    ),
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}
