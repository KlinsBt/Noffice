import PptxGenJS from 'pptxgenjs';
import { quadrantPicture } from './pptx-retained';

export async function stackPresentationFixture(): Promise<ArrayBuffer> {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  const slide = pptx.addSlide();
  slide.addText('Bottom text', {
    objectName: 'Bottom text',
    x: 1,
    y: 1,
    w: 4,
    h: 2,
    fill: { color: 'DDDDDD' },
    fontSize: 24,
  });
  slide.addImage({
    objectName: 'Middle image',
    data: `data:image/png;base64,${quadrantPicture}`,
    x: 2,
    y: 1.5,
    w: 3,
    h: 2,
  });
  slide.addChart(pptx.ChartType.bar, [{ name: 'Data', labels: ['One', 'Two'], values: [2, 5] }], {
    objectName: 'Retained chart',
    x: 8,
    y: 1,
    w: 4,
    h: 3,
  });
  slide.addText('Top text', {
    objectName: 'Top text',
    x: 3,
    y: 2,
    w: 4,
    h: 2,
    fontSize: 24,
    fill: { color: 'CCCCFF' },
  });
  slide.addNotes('Layering retains notes');
  pptx.addSlide().addText('Untouched slide', { x: 1, y: 1, w: 5, h: 1 });
  return (await pptx.write({ outputType: 'arraybuffer' })) as ArrayBuffer;
}
