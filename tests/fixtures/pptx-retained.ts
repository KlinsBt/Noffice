import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
export const pictureData =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l8sAAAAASUVORK5CYII=';
// Authored 2x2 red/green/blue/yellow pixels make mirror direction observable.
export const quadrantPicture =
  'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFElEQVR4nGP4z8DAAMIM////ZwAAHu8E/KPItPcAAAAASUVORK5CYII=';
export async function flippedImageFixture(): Promise<ArrayBuffer> {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'FLIP_4_3', width: 10, height: 7.5 });
  pptx.layout = 'FLIP_4_3';
  const slide = pptx.addSlide();
  slide.addImage({
    data: `data:image/png;base64,${quadrantPicture}`,
    x: 2,
    y: 2,
    w: 3,
    h: 2,
    flipH: true,
  });
  slide.addText('Flip reference', { x: 1, y: 0.5, w: 5, h: 1 });
  return (await pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
}
export async function retainedPresentationFixture(withNotes = true): Promise<ArrayBuffer> {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'TEST_4_3', width: 10, height: 7.5 });
  pptx.layout = 'TEST_4_3';
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  slide.addText(
    [
      { text: 'Original ', options: { bold: true, color: 'B00020', fontFace: 'Arial' } },
      { text: 'title', options: { italic: true, color: '0033AA', fontFace: 'Georgia' } },
    ],
    { x: 1, y: 1, w: 6, h: 1, fontSize: 24, breakLine: false, margin: 0 },
  );
  slide.addShape(pptx.ShapeType.rect, {
    x: 1,
    y: 3,
    w: 2,
    h: 1,
    fill: { color: '22AA44' },
    line: { color: '111111', width: 2 },
  });
  slide.addNotes('Original speaker notes');
  const chart = pptx.addSlide();
  chart.addChart(
    pptx.ChartType.bar,
    [{ name: 'Series', labels: ['Alpha', 'Beta'], values: [3, 7] }],
    { x: 1, y: 1, w: 6, h: 4 },
  );
  chart.addNotes('Chart notes');
  const output = (await pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
  if (withNotes) return output;
  const zip = await JSZip.loadAsync(output);
  for (const path of Object.keys(zip.files)) {
    if (path.startsWith('ppt/notesSlides/') || path.startsWith('ppt/notesMasters/')) {
      zip.remove(path);
      continue;
    }
    if (
      path.endsWith('.rels') ||
      path === '[Content_Types].xml' ||
      path === 'ppt/presentation.xml'
    ) {
      let xml = await zip.file(path)!.async('string');
      xml = xml
        .replace(
          /<Relationship\b[^>]*Type="[^"]*\/notes(?:Slide|Master)"[^>]*\/?>(?:<\/Relationship>)?/g,
          '',
        )
        .replace(
          /<Override\b[^>]*PartName="\/ppt\/notes(?:Slides|Masters)\/[^"]+"[^>]*\/?>(?:<\/Override>)?/g,
          '',
        )
        .replace(/<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>/g, '');
      zip.file(path, xml, { createFolders: false });
    }
  }
  return zip.generateAsync({ type: 'arraybuffer' });
}
